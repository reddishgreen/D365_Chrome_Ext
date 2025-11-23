import { CrmApi } from '../utils/api';
import { 
    EntityMetadata, 
    AttributeMetadata, 
    RelationshipMetadata, 
    QueryFilter, 
    QueryColumn, 
    JoinedEntity 
} from '../types';

export interface ParseResult {
    filters: QueryFilter;
    columns: QueryColumn[];
    joinedEntities: JoinedEntity[];
}

// Helper to map FetchXML operators to our internal operators
export const mapFetchXmlOperator = (op: string): string => {
    switch (op) {
        case 'eq': return 'eq';
        case 'ne': return 'ne';
        case 'like': return 'contains'; // We will need to strip %
        case 'not-like': return 'not contains';
        case 'begins-with': return 'startswith';
        case 'ends-with': return 'endswith';
        case 'gt': return 'gt';
        case 'ge': return 'ge';
        case 'lt': return 'lt';
        case 'le': return 'le';
        case 'null': return 'null';
        case 'not-null': return 'not null';
        case 'today': return 'today';
        case 'yesterday': return 'yesterday';
        case 'tomorrow': return 'tomorrow';
        case 'this-week': return 'this-week';
        case 'last-week': return 'last-week';
        case 'next-week': return 'next-week';
        case 'last-x-days': return 'last-x-days';
        case 'next-x-days': return 'next-x-days';
        case 'last-x-months': return 'last-x-months';
        case 'next-x-months': return 'next-x-months';
        case 'last-x-years': return 'last-x-years';
        // Add more as needed
        default: return 'eq'; // Fallback
    }
};

export const parseValue = (val: string | null, operator: string, meta?: AttributeMetadata): any => {
    if (val === null || val === undefined) return null;
    
    // Handle wildcard stripping for 'like' mapped to 'contains'
    if (operator === 'contains' || operator === 'not contains') {
        return val.replace(/%/g, '');
    }

    // If the operator expects a number (e.g. last-x-days), try to parse it as a number
    // even if the attribute type is DateTime.
    if (operator.includes('x-')) {
        const num = Number(val);
        return isNaN(num) ? val : num;
    }

    if (!meta) return val;

    const type = meta.AttributeType;
    if (type === 'Integer' || type === 'Money' || type === 'Decimal' || type === 'Double' || type === 'Picklist' || type === 'State' || type === 'Status') {
        const num = Number(val);
        return isNaN(num) ? val : num;
    }
    if (type === 'Boolean') {
        return val === '1' || val.toLowerCase() === 'true';
    }
    
    return val;
};

export const parseFilterNode = (node: Element, alias: string, attrMap: Map<string, AttributeMetadata>): QueryFilter | null => {
    const type = node.getAttribute('type') || 'and';
    
    // It's a group
    const filterGroup: QueryFilter = {
        id: Math.random().toString(36).substr(2, 9),
        type: 'group',
        logicalOperator: type === 'or' ? 'or' : 'and',
        children: []
    };

    const childNodes = Array.from(node.children);
    for (const child of childNodes) {
        if (child.nodeName === 'filter') {
            const subGroup = parseFilterNode(child, alias, attrMap);
            if (subGroup) filterGroup.children?.push(subGroup);
        } else if (child.nodeName === 'condition') {
            const attrName = child.getAttribute('attribute');
            const op = child.getAttribute('operator');
            const val = child.getAttribute('value');
            
            if (attrName && op) {
                const meta = attrMap.get(attrName.toLowerCase());
                const mappedOp = mapFetchXmlOperator(op);
                const parsedVal = parseValue(val, mappedOp, meta);
                
                filterGroup.children?.push({
                    id: Math.random().toString(36).substr(2, 9),
                    type: 'condition',
                    entityAlias: alias,
                    attribute: attrName,
                    operator: mappedOp,
                    value: parsedVal
                });
            }
        }
    }
    
    // Only return if it has children or we want to preserve empty groups (rare)
    if (filterGroup.children && filterGroup.children.length > 0) {
        return filterGroup;
    }
    return null;
};

export const parseViewFetchXml = async (
    fetchXml: string, 
    api: CrmApi, 
    mainEntity: { logicalName: string, primaryIdAttribute: string, primaryNameAttribute: string }
): Promise<ParseResult> => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(fetchXml, 'application/xml');
    
    const entityNode = doc.querySelector('fetch > entity');
    if (!entityNode) {
        throw new Error('Invalid FetchXML: No root entity node found');
    }

    const attributeCache = new Map<string, AttributeMetadata[]>();
    const entityCache = new Map<string, EntityMetadata>();
    const relationshipCache = new Map<string, {
        manyToOne: RelationshipMetadata[];
        oneToMany: RelationshipMetadata[];
    }>();

    const getAttributesFor = async (logicalName: string): Promise<AttributeMetadata[]> => {
        const key = logicalName.toLowerCase();
        if (attributeCache.has(key)) {
            return attributeCache.get(key)!;
        }
        const attrs = await api.getAttributes(logicalName);
        attributeCache.set(key, attrs);
        return attrs;
    };

    const getEntityMetadataFor = async (logicalName: string): Promise<EntityMetadata> => {
        const key = logicalName.toLowerCase();
        if (entityCache.has(key)) {
            return entityCache.get(key)!;
        }
        const metadata = await api.getEntity(logicalName);
        entityCache.set(key, metadata);
        return metadata;
    };

    const getRelationshipsFor = async (logicalName: string) => {
        const key = logicalName.toLowerCase();
        if (relationshipCache.has(key)) {
            return relationshipCache.get(key)!;
        }
        const manyToOne = await api.getRelationships(logicalName, 'ManyToOne');
        const oneToMany = await api.getRelationships(logicalName, 'OneToMany');
        const value = { manyToOne, oneToMany };
        relationshipCache.set(key, value);
        return value;
    };

    const resolveNavigation = async (
        parentLogical: string,
        childLogical: string,
        fromAttr?: string | null,
        toAttr?: string | null
    ): Promise<{ relationship: RelationshipMetadata; navigationProperty: string } | null> => {
        const rels = await getRelationshipsFor(parentLogical);
        const childLower = childLogical.toLowerCase();
        const toLower = toAttr ? toAttr.toLowerCase() : undefined;
        const fromLower = fromAttr ? fromAttr.toLowerCase() : undefined;

        const manyMatch = rels.manyToOne.find(r =>
            r.ReferencedEntity.toLowerCase() === childLower &&
            (!!toLower ? r.ReferencingAttribute.toLowerCase() === toLower : true)
        );
        if (manyMatch && manyMatch.ReferencingEntityNavigationPropertyName) {
            return { relationship: manyMatch, navigationProperty: manyMatch.ReferencingEntityNavigationPropertyName };
        }

        const oneMatch = rels.oneToMany.find(r =>
            r.ReferencingEntity.toLowerCase() === childLower &&
            (!!fromLower ? r.ReferencingAttribute.toLowerCase() === fromLower : true) &&
            r.ReferencedEntityNavigationPropertyName
        );
        if (oneMatch && oneMatch.ReferencedEntityNavigationPropertyName) {
            return { relationship: oneMatch, navigationProperty: oneMatch.ReferencedEntityNavigationPropertyName };
        }

        return null;
    };

    // Fetch metadata for main entity
    const attributesMeta = await getAttributesFor(mainEntity.logicalName);

    const validAttributesMap = new Map<string, AttributeMetadata>();
    attributesMeta.forEach(attr => {
        validAttributesMap.set(attr.LogicalName.toLowerCase(), attr);
    });

    const newColumns: QueryColumn[] = [];
    const newJoinedEntities: JoinedEntity[] = [];
    const newFilterChildren: QueryFilter[] = [];

    const pushColumn = (col: QueryColumn) => {
        const key = `${col.entityAlias}:${col.attribute}`;
        if (!newColumns.some(c => `${c.entityAlias}:${c.attribute}` === key)) {
            newColumns.push(col);
        }
    };

    const getSelectAttributeName = (meta: AttributeMetadata): string => {
        const type = meta.AttributeType?.toLowerCase();
        if (type === 'lookup' || type === 'customer' || type === 'owner') {
            return `_${meta.LogicalName}_value`;
        }
        return meta.LogicalName;
    };

    const pushMetaColumn = (entityAlias: string, meta: AttributeMetadata, displayOverride?: string) => {
        pushColumn({
            entityAlias,
            attribute: getSelectAttributeName(meta),
            displayName: displayOverride || meta.DisplayName,
            logicalName: meta.LogicalName,
            attributeType: meta.AttributeType
        });
    };

    // Add main entity primary ID
    const primaryIdMeta = attributesMeta.find(a => a.LogicalName === mainEntity.primaryIdAttribute);
    if (primaryIdMeta) {
        pushMetaColumn('main', primaryIdMeta);
    }

    // Get attributes from main entity only (not from link-entity)
    const attributeNodes = Array.from(entityNode.children).filter(node => node.nodeName === 'attribute');
    attributeNodes.forEach(node => {
        const attrName = node.getAttribute('name');
        const alias = node.getAttribute('alias');
        if (!attrName) return;
        
        const meta = validAttributesMap.get(attrName.toLowerCase());
        if (meta) {
            pushMetaColumn('main', meta, alias || meta.DisplayName);
        }
    });

    // Parse Main Entity Filters
    const mainFilterNodes = Array.from(entityNode.children).filter(node => node.nodeName === 'filter');
    mainFilterNodes.forEach(node => {
        const filter = parseFilterNode(node, 'main', validAttributesMap);
        if (filter) newFilterChildren.push(filter);
    });

    // If no columns selected, add primary name as fallback
    if (newColumns.length === 0) {
        const nameMeta = attributesMeta.find(a => a.LogicalName === mainEntity.primaryNameAttribute);
        if (nameMeta) {
            pushMetaColumn('main', nameMeta);
        }
    }

    const processLinkEntities = async (
        container: Element,
        parentAlias: string,
        parentLogical: string
    ): Promise<void> => {
        const linkNodes = Array.from(container.children).filter(node => node.nodeName === 'link-entity');
        for (const node of linkNodes) {
            const link = node as Element;
            const childLogical = link.getAttribute('name');
            if (!childLogical) continue;
            const aliasAttr = link.getAttribute('alias');
            const alias = aliasAttr || childLogical;
            const fromAttr = link.getAttribute('from');
            const toAttr = link.getAttribute('to');

            const navInfo = await resolveNavigation(parentLogical, childLogical, fromAttr, toAttr);
            if (!navInfo) {
                console.warn(`Unable to resolve relationship for link-entity '${alias}' (${parentLogical} -> ${childLogical}). Skipping.`);
                continue;
            }

            const childMetadata = await getEntityMetadataFor(childLogical);
            const childAttributes = await getAttributesFor(childLogical);
            const childAttrMap = new Map<string, AttributeMetadata>();
            childAttributes.forEach(attr => childAttrMap.set(attr.LogicalName.toLowerCase(), attr));

            const joinedAliasKey = alias.toLowerCase();
            if (!newJoinedEntities.some(j => j.alias.toLowerCase() === joinedAliasKey && j.parentAlias === parentAlias)) {
                newJoinedEntities.push({
                    logicalName: childMetadata.LogicalName,
                    entitySetName: childMetadata.EntitySetName,
                    displayName: childMetadata.DisplayName,
                    primaryIdAttribute: childMetadata.PrimaryIdAttribute,
                    primaryNameAttribute: childMetadata.PrimaryNameAttribute,
                    alias,
                    relationshipName: navInfo.relationship.SchemaName,
                    parentAlias,
                    relationshipType: navInfo.relationship.RelationshipType,
                    navigationPropertyName: navInfo.navigationProperty
                });
            }

            // Collect columns for this linked entity
            const childAttributeNodes = Array.from(link.children).filter(n => n.nodeName === 'attribute');
            for (const attrNode of childAttributeNodes) {
                const attrName = attrNode.getAttribute('name');
                if (!attrName) continue;
                const childMeta = childAttrMap.get(attrName.toLowerCase());
                if (!childMeta) continue;
                const attrAlias = attrNode.getAttribute('alias');
                pushMetaColumn(alias, childMeta, attrAlias || childMeta.DisplayName);
            }

            // Collect filters for this linked entity
            const childFilterNodes = Array.from(link.children).filter(n => n.nodeName === 'filter');
            childFilterNodes.forEach(node => {
                const filter = parseFilterNode(node, alias, childAttrMap);
                if (filter) newFilterChildren.push(filter);
            });

            // Recurse into nested link-entities
            await processLinkEntities(link, alias, childLogical);
        }
    };

    await processLinkEntities(entityNode, 'main', mainEntity.logicalName);

    return {
        filters: {
            id: 'root',
            type: 'group',
            logicalOperator: 'and',
            children: newFilterChildren
        },
        columns: newColumns,
        joinedEntities: newJoinedEntities
    };
};
