import { EntityMetadataComplete, AttributeMetadataComplete, RelationshipMetadata } from '../../query-builder/types';

export type VerbosityLevel = 'compact' | 'standard' | 'full';

export interface PromptOptions {
  verbosity: VerbosityLevel;
  includeRules: boolean;
  orgUrl?: string;
  apiVersion?: string; // e.g., 'v9.2'
}

// API version constant - matches CrmApi default
const DEFAULT_API_VERSION = 'v9.2';

// Helper to extract base URL from org URL
function getOrgBaseUrl(orgUrl?: string): string | null {
  if (!orgUrl) return null;
  try {
    const url = new URL(orgUrl);
    return `${url.protocol}//${url.host}`;
  } catch {
    return null;
  }
}

// Helper to get API version
function getApiVersion(options: PromptOptions): string {
  return options.apiVersion || DEFAULT_API_VERSION;
}

// Helper to build example queries for an entity
function buildExampleQueries(
  entityMeta: EntityMetadataComplete,
  apiVersion: string,
  baseUrl: string | null
): string[] {
  const lines: string[] = [];
  const entitySet = entityMeta.EntitySetName;
  const primaryId = entityMeta.PrimaryIdAttribute;
  const primaryName = entityMeta.PrimaryNameAttribute || primaryId;
  const apiPath = `/api/data/${apiVersion}/${entitySet}`;

  lines.push('### Example Queries');
  lines.push('');

  // Get top 10
  lines.push(`- **Get top 10 records**:`);
  lines.push(`  \`GET ${apiPath}?$select=${primaryId},${primaryName}&$top=10\``);
  lines.push('');

  // Get by ID
  lines.push(`- **Get by ID**:`);
  lines.push(`  \`GET ${apiPath}(<guid>)?$select=${primaryId},${primaryName}\``);
  lines.push('');

  // Activity pointer special case
  if (entityMeta.LogicalName === 'activitypointer' || entityMeta.EntitySetName === 'activitypointers') {
    lines.push(`- **Get activities for a record** (using regarding filter):`);
    lines.push(`  \`GET /api/data/${apiVersion}/activitypointers?$select=activityid,subject,regardingobjectid&$filter=_regardingobjectid_value eq <guid>&$top=10\``);
    lines.push('');
  }

  return lines;
}

export interface EntitySelection {
  entityLogicalName: string;
  selectedFields: string[]; // Logical names of selected fields
  selectedRelationships: string[]; // Schema names of selected relationships
}

export interface PromptSelections {
  entities: Map<string, EntitySelection>;
  selectedRelationships: string[]; // Global relationship selections
}

export function generatePromptMarkdown(
  metadataMap: Map<string, EntityMetadataComplete>,
  selections: PromptSelections,
  options: PromptOptions
): string {
  const lines: string[] = [];

  // Header section
  if (options.includeRules) {
    lines.push('# Dynamics 365 Entity Context');
    lines.push('');
    lines.push('## Context Information');
    if (options.orgUrl) {
      const orgName = new URL(options.orgUrl).hostname;
      lines.push(`- **Environment**: ${orgName}`);
    }
    lines.push(`- **Generated**: ${new Date().toISOString()}`);
    lines.push('');

    // Dataverse Web API Reference section
    const baseUrl = getOrgBaseUrl(options.orgUrl);
    const apiVersion = getApiVersion(options);
    
    lines.push('## Dataverse Web API Reference');
    lines.push('');
    if (baseUrl) {
      lines.push(`- **Base URL**: \`${baseUrl}\``);
    }
    lines.push(`- **Web API Root**: \`/api/data/${apiVersion}/\``);
    
    // Example entity set URL using first selected entity
    const firstEntity = Array.from(selections.entities.keys())[0];
    if (firstEntity) {
      const firstEntityMeta = metadataMap.get(firstEntity);
      if (firstEntityMeta && baseUrl) {
        lines.push(`- **Example Entity Set URL**: \`${baseUrl}/api/data/${apiVersion}/${firstEntityMeta.EntitySetName}\``);
      }
    }
    lines.push('');
    lines.push('**Naming + Casing Rules**:');
    lines.push('- Use entity set names exactly as returned by metadata');
    lines.push('- Use attribute logical names exactly as returned by metadata');
    lines.push('- Use standard OData query option casing: `$select`, `$filter`, `$expand`, `$orderby`, `$top`, `$count`');
    lines.push('- Don\'t invent casing / don\'t title-case fields');
    lines.push('');

    lines.push('## Quick Reference');
    const entityNames: string[] = [];
    selections.entities.forEach((selection, logicalName) => {
      const meta = metadataMap.get(logicalName);
      if (meta) {
        entityNames.push(`- **${meta.DisplayName}** (Logical: \`${meta.LogicalName}\`, Entity Set: \`${meta.EntitySetName}\`)`);
      }
    });
    lines.push(...entityNames);
    lines.push('');
  }

  // Entity details
  selections.entities.forEach((selection, entityLogicalName) => {
    const entityMeta = metadataMap.get(entityLogicalName);
    if (!entityMeta) return;

    lines.push(`## ${entityMeta.DisplayName} Entity Details`);
    lines.push('');
    lines.push(`**Logical Name**: \`${entityMeta.LogicalName}\``);
    lines.push(`**Entity Set**: \`${entityMeta.EntitySetName}\``);
    lines.push(`**Primary ID**: \`${entityMeta.PrimaryIdAttribute}\``);
    lines.push(`**Primary Name**: \`${entityMeta.PrimaryNameAttribute}\``);
    if (entityMeta.Description) {
      lines.push(`**Description**: ${entityMeta.Description}`);
    }
    lines.push('');

    // Add example queries
    const baseUrl = getOrgBaseUrl(options.orgUrl);
    const apiVersion = getApiVersion(options);
    const exampleQueries = buildExampleQueries(entityMeta, apiVersion, baseUrl);
    lines.push(...exampleQueries);

    // Selected fields
    const selectedFields = entityMeta.Attributes.filter(attr =>
      selection.selectedFields.includes(attr.LogicalName)
    );

    if (selectedFields.length > 0) {
      lines.push('### Selected Fields');
      lines.push('');

      selectedFields.forEach(attr => {
        lines.push(`- \`${attr.LogicalName}\` (Display: ${attr.DisplayName})`);

        if (options.verbosity === 'compact') {
          lines.push(`  - Type: ${attr.AttributeType}`);
          if (attr.RequiredLevel === 'ApplicationRequired' || attr.RequiredLevel === 'SystemRequired') {
            lines.push(`  - Required: Yes`);
          }
          if (attr.LookupTargets && attr.LookupTargets.length > 0) {
            lines.push(`  - Lookup Targets: ${attr.LookupTargets.join(', ')}`);
          }
          // Option sets (compact - first 10)
          if (attr.OptionSetValues && attr.OptionSetValues.length > 0) {
            const displayValues = attr.OptionSetValues.slice(0, 10).map(opt =>
              `${opt.Value} = ${opt.Label}`
            ).join(', ');
            if (attr.OptionSetValues.length > 10) {
              lines.push(`  - Options: ${displayValues}... (+${attr.OptionSetValues.length - 10} more)`);
            } else {
              lines.push(`  - Options: ${displayValues}`);
            }
          }
        } else if (options.verbosity === 'standard') {
          lines.push(`  - Type: ${attr.AttributeType}`);
          if (attr.RequiredLevel === 'ApplicationRequired' || attr.RequiredLevel === 'SystemRequired') {
            lines.push(`  - Required: Yes`);
          } else if (attr.RequiredLevel === 'Recommended') {
            lines.push(`  - Required: Recommended`);
          } else {
            lines.push(`  - Required: No`);
          }
          
          if (attr.MaxLength !== null && attr.MaxLength !== undefined) {
            lines.push(`  - Max Length: ${attr.MaxLength}`);
          }
          if (attr.Precision !== null && attr.Precision !== undefined) {
            lines.push(`  - Precision: ${attr.Precision}`);
          }
          if (attr.Scale !== null && attr.Scale !== undefined) {
            lines.push(`  - Scale: ${attr.Scale}`);
          }
          
          if (attr.Description) {
            lines.push(`  - Description: ${attr.Description}`);
          }

          // Option sets (truncated to 50)
          if (attr.OptionSetValues && attr.OptionSetValues.length > 0) {
            const displayValues = attr.OptionSetValues.slice(0, 50).map(opt =>
              `${opt.Value} = ${opt.Label}`
            ).join(', ');
            if (attr.OptionSetValues.length > 50) {
              lines.push(`  - Options (first 50 of ${attr.OptionSetValues.length}): ${displayValues}...`);
            } else {
              lines.push(`  - Options: ${displayValues}`);
            }
          }

          if (attr.LookupTargets && attr.LookupTargets.length > 0) {
            lines.push(`  - Lookup Targets: ${attr.LookupTargets.join(', ')}`);
            if (attr.IsPolymorphic) {
              lines.push(`  - Note: Polymorphic lookup (Customer/Regarding type)`);
            }
          }

          if (attr.IsCalculated) {
            lines.push(`  - Note: Calculated field`);
          }

          if (attr.IsRollup) {
            lines.push(`  - Note: Rollup field`);
          }
        } else {
          // Full verbosity
          lines.push(`  - Type: ${attr.AttributeType}`);
          lines.push(`  - Required Level: ${attr.RequiredLevel || 'None'}`);
          if (attr.MaxLength !== null && attr.MaxLength !== undefined) {
            lines.push(`  - Max Length: ${attr.MaxLength}`);
          }
          if (attr.Precision !== null && attr.Precision !== undefined) {
            lines.push(`  - Precision: ${attr.Precision}`);
          }
          if (attr.Scale !== null && attr.Scale !== undefined) {
            lines.push(`  - Scale: ${attr.Scale}`);
          }
          if (attr.Description) {
            lines.push(`  - Description: ${attr.Description}`);
          }
          if (attr.IsPrimaryId) {
            lines.push(`  - Primary Key: Yes`);
          }
          if (attr.IsPrimaryName) {
            lines.push(`  - Primary Name: Yes`);
          }

          // Full option set values
          if (attr.OptionSetValues && attr.OptionSetValues.length > 0) {
            lines.push(`  - Options (${attr.OptionSetValues.length}):`);
            attr.OptionSetValues.forEach(opt => {
              lines.push(`    - ${opt.Value} = ${opt.Label}`);
            });
          }

          if (attr.LookupTargets && attr.LookupTargets.length > 0) {
            lines.push(`  - Lookup Targets: ${attr.LookupTargets.join(', ')}`);
            if (attr.IsPolymorphic) {
              lines.push(`  - Polymorphic: Yes`);
            }
          }

          if (attr.IsCalculated) {
            lines.push(`  - Calculated: Yes`);
          }

          if (attr.IsRollup) {
            lines.push(`  - Rollup: Yes`);
          }
        }

        lines.push('');
      });
    }

    lines.push('');
  });

  // Relationships section
  const allRelationships: Array<{ entity: string; relationship: RelationshipMetadata }> = [];
  
  selections.entities.forEach((selection, entityLogicalName) => {
    const entityMeta = metadataMap.get(entityLogicalName);
    if (!entityMeta) return;

    // Collect all relationships for selected entities
    [...entityMeta.OneToManyRelationships, ...entityMeta.ManyToOneRelationships, ...entityMeta.ManyToManyRelationships]
      .filter(rel => selection.selectedRelationships.includes(rel.SchemaName))
      .forEach(rel => {
        allRelationships.push({ entity: entityLogicalName, relationship: rel });
      });
  });

  // Also include globally selected relationships
  selections.selectedRelationships.forEach(relSchemaName => {
    metadataMap.forEach((entityMeta, entityLogicalName) => {
      const allRels = [
        ...entityMeta.OneToManyRelationships,
        ...entityMeta.ManyToOneRelationships,
        ...entityMeta.ManyToManyRelationships
      ];
      const rel = allRels.find(r => r.SchemaName === relSchemaName);
      if (rel && !allRelationships.some(ar => ar.relationship.SchemaName === relSchemaName)) {
        allRelationships.push({ entity: entityLogicalName, relationship: rel });
      }
    });
  });

  if (allRelationships.length > 0) {
    lines.push('## Relationships');
    lines.push('');

    // Group by relationship type
    const byType: Record<string, typeof allRelationships> = {
      OneToMany: [],
      ManyToOne: [],
      ManyToMany: []
    };

    allRelationships.forEach(({ entity, relationship }) => {
      byType[relationship.RelationshipType].push({ entity, relationship });
    });

    Object.entries(byType).forEach(([type, rels]) => {
      if (rels.length === 0) return;

      lines.push(`### ${type} Relationships`);
      lines.push('');

      rels.forEach(({ entity, relationship }) => {
        const sourceEntity = metadataMap.get(relationship.ReferencingEntity);
        const targetEntity = metadataMap.get(relationship.ReferencedEntity);
        const sourceName = sourceEntity?.DisplayName || relationship.ReferencingEntity;
        const targetName = targetEntity?.DisplayName || relationship.ReferencedEntity;
        const sourceEntitySet = sourceEntity?.EntitySetName || relationship.ReferencingEntity;
        const targetEntitySet = targetEntity?.EntitySetName || relationship.ReferencedEntity;
        const isSelfReferencing = relationship.ReferencingEntity === relationship.ReferencedEntity;

        lines.push(`- **${relationship.SchemaName}**`);
        if (isSelfReferencing) {
          lines.push(`  - Self-referencing relationship: ${sourceName} → ${sourceName}`);
        } else {
          lines.push(`  - ${sourceName} → ${targetName}`);
        }
        lines.push(`  - Schema Name: \`${relationship.SchemaName}\``);
        lines.push(`  - Type: ${relationship.RelationshipType}`);
        lines.push(`  - Source Entity Set: \`${sourceEntitySet}\``);
        lines.push(`  - Target Entity Set: \`${targetEntitySet}\``);

        // Lookup Attribute (referencing attribute) - this is NOT a navigation property
        if (relationship.ReferencingAttribute) {
          lines.push(`  - Lookup Attribute (on ${sourceName}): \`${relationship.ReferencingAttribute}\``);
        }

        // Navigation properties with $expand examples
        if (relationship.RelationshipType === 'OneToMany' && relationship.ReferencingEntityNavigationPropertyName) {
          lines.push(`  - Collection Navigation Property (1:N): \`${relationship.ReferencingEntityNavigationPropertyName}\``);
          lines.push(`  - Example \`$expand\`: \`/${sourceEntitySet}?$expand=${relationship.ReferencingEntityNavigationPropertyName}($select=...)\``);
        } else if (relationship.RelationshipType === 'ManyToOne' && relationship.ReferencedEntityNavigationPropertyName) {
          lines.push(`  - Single-valued Navigation Property (N:1): \`${relationship.ReferencedEntityNavigationPropertyName}\``);
          lines.push(`  - Example \`$expand\`: \`/${sourceEntitySet}?$expand=${relationship.ReferencedEntityNavigationPropertyName}($select=...)\``);
        } else if (relationship.RelationshipType === 'ManyToMany') {
          if (relationship.ReferencingEntityNavigationPropertyName) {
            lines.push(`  - Collection Navigation Property (M:N, on ${sourceName}): \`${relationship.ReferencingEntityNavigationPropertyName}\``);
            lines.push(`  - Example \`$expand\`: \`/${sourceEntitySet}?$expand=${relationship.ReferencingEntityNavigationPropertyName}($select=...)\``);
          }
          if (relationship.ReferencedEntityNavigationPropertyName) {
            lines.push(`  - Collection Navigation Property (M:N, on ${targetName}): \`${relationship.ReferencedEntityNavigationPropertyName}\``);
            lines.push(`  - Example \`$expand\`: \`/${targetEntitySet}?$expand=${relationship.ReferencedEntityNavigationPropertyName}($select=...)\``);
          }
        }

        lines.push('');
      });
    });
  }

  return lines.join('\n');
}
