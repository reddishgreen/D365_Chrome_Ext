import { AttributeMetadataComplete, EntityMetadataComplete, RelationshipMetadata } from '../../query-builder/types';

export type VerbosityLevel = 'compact' | 'standard' | 'full';

export interface PromptOptions {
  verbosity: VerbosityLevel;
  includeRules: boolean;
  orgUrl?: string;
  apiVersion?: string; // e.g., 'v9.2'
  generatedAt?: string; // ISO timestamp; if omitted, header is omitted
}

const DEFAULT_API_VERSION = 'v9.2';

function getOrgBaseUrl(orgUrl?: string): string | null {
  if (!orgUrl) return null;
  try {
    const url = new URL(orgUrl);
    return `${url.protocol}//${url.host}`;
  } catch {
    return null;
  }
}

function getApiVersion(options: PromptOptions): string {
  return options.apiVersion || DEFAULT_API_VERSION;
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

/** Lookup-style attribute types — use `_<logical>_value` for $select/$filter. */
const LOOKUP_TYPES = new Set(['Lookup', 'Customer', 'Owner']);

/** Escape pipe and newlines so a value is safe inside a markdown table cell. */
function tcell(value: string | number | undefined | null): string {
  if (value === undefined || value === null) return '';
  return String(value).replace(/\|/g, '\\|').replace(/\r?\n/g, ' ');
}

function typeWithSize(attr: AttributeMetadataComplete): string {
  if (attr.AttributeType === 'String' || attr.AttributeType === 'Memo') {
    return attr.MaxLength != null ? `${attr.AttributeType}(${attr.MaxLength})` : attr.AttributeType;
  }
  if ((attr.AttributeType === 'Decimal' || attr.AttributeType === 'Money') && attr.Precision != null) {
    return `${attr.AttributeType}(p=${attr.Precision})`;
  }
  return attr.AttributeType;
}

function requiredShort(level?: string): string {
  switch (level) {
    case 'SystemRequired': return 'System';
    case 'ApplicationRequired': return 'Required';
    case 'Recommended': return 'Recommended';
    default: return '';
  }
}

function fieldNotes(attr: AttributeMetadataComplete): string {
  const notes: string[] = [];
  if (attr.IsPrimaryId) notes.push('PK');
  if (attr.IsPrimaryName) notes.push('Primary Name');
  if (attr.AttributeOf) notes.push(`shadow of \`${attr.AttributeOf}\``);
  if (attr.IsCalculated) notes.push('calculated');
  if (attr.IsRollup) notes.push('rollup');
  // settability — only flag when explicitly false; otherwise assume true
  if (attr.IsValidForCreate === false) notes.push('not settable on create');
  if (attr.IsValidForUpdate === false) notes.push('not settable on update');
  if (attr.IsValidForRead === false) notes.push('not readable');
  return notes.join('; ');
}

function isReadOnly(attr: AttributeMetadataComplete): boolean {
  return attr.IsValidForCreate === false && attr.IsValidForUpdate === false;
}

function buildCheatSheet(): string[] {
  return [
    '## OData Quick Reference',
    '',
    'Critical syntax for Dataverse Web API queries — apply consistently:',
    '',
    '- **Lookup field GUIDs**: use `_<logical>_value` in `$select` and `$filter`. The bare logical name is not a queryable column for lookups; only the `_value` form returns the GUID.',
    '  - Example: `?$select=_primarycontactid_value&$filter=_primarycontactid_value eq <guid>`',
    '- **Lookup expand**: use the **navigation property name** (not the logical name).',
    '  - Example: `?$expand=primarycontactid_contact($select=fullname)`',
    '- **Bind on create / update**: `"<lookup>@odata.bind": "/<targetEntitySet>(<guid>)"`',
    '- **Polymorphic lookups** (Customer, regarding, owner): include the target entity in the property name.',
    '  - Example: `"regardingobjectid_account@odata.bind": "/accounts(<guid>)"`',
    '- **Datetime filters**: ISO-8601 with `Z`, unquoted: `createdon ge 2025-01-01T00:00:00Z`',
    '- **GUID filters**: unquoted: `accountid eq 00000000-0000-0000-0000-000000000000`',
    '- **String filters**: single-quoted, escape `\'` by doubling: `name eq \'O\'\'Brien\'`',
    '- **Option-set / state / status filters**: filter by **integer value**, not label.',
    '- **Boolean filters**: lowercase `true` / `false`.',
    '- **Headers**:',
    '  - `Prefer: return=representation` — return the record body on Create/Update',
    '  - `Prefer: odata.include-annotations="*"` — include `FormattedValue` and `OData.Community.Display.V1.FormattedValue` annotations (lookup display names, option-set labels, formatted dates)',
    '  - `If-Match: *` for conditional update; `If-None-Match: *` for upsert-as-create',
    '- **Counting**: `?$count=true` adds `@odata.count`. `?$apply=aggregate(...)` for grouping.',
    '- **Read-only fields** (e.g. `createdon`, `modifiedby`, lookup-name shadow attributes) are populated by the platform — never include them in Create/Update bodies.',
    '',
  ];
}

function buildFieldTable(
  fields: AttributeMetadataComplete[],
  verbosity: VerbosityLevel
): string[] {
  if (fields.length === 0) return [];
  const lines: string[] = [];
  lines.push('| Logical Name | Display | Type | Required | Notes |');
  lines.push('|---|---|---|---|---|');
  fields.forEach(attr => {
    const notes: string[] = [];
    const baseNotes = fieldNotes(attr);
    if (baseNotes) notes.push(baseNotes);
    if (verbosity === 'full' && attr.Description) {
      notes.push(attr.Description);
    }
    lines.push(
      `| \`${tcell(attr.LogicalName)}\` | ${tcell(attr.DisplayName)} | ${tcell(typeWithSize(attr))} | ${tcell(requiredShort(attr.RequiredLevel))} | ${tcell(notes.join('; '))} |`
    );
  });
  lines.push('');
  return lines;
}

function buildLookupTable(
  fields: AttributeMetadataComplete[],
  targetEntitySetMap: Map<string, string>
): string[] {
  const lookups = fields.filter(f => LOOKUP_TYPES.has(f.AttributeType));
  if (lookups.length === 0) return [];

  const lines: string[] = [];
  lines.push('### Lookup Fields');
  lines.push('');
  lines.push('Use `_<logical>_value` for `$select` / `$filter`. Use `@odata.bind` for write operations.');
  lines.push('');
  lines.push('| Logical | $select / $filter | Targets | @odata.bind example |');
  lines.push('|---|---|---|---|');
  lookups.forEach(attr => {
    const targets = attr.LookupTargets || [];
    const bind = targets.length === 0
      ? `"${attr.LogicalName}@odata.bind": "/<entityset>(<guid>)"`
      : targets.length === 1
        ? `"${attr.LogicalName}@odata.bind": "/${targetEntitySetMap.get(targets[0]) || targets[0]}(<guid>)"`
        : targets
            .map(t => `"${attr.LogicalName}_${t}@odata.bind": "/${targetEntitySetMap.get(t) || t}(<guid>)"`)
            .join(' OR ');
    lines.push(
      `| \`${tcell(attr.LogicalName)}\` | \`_${attr.LogicalName}_value\` | ${tcell(targets.join(', '))}${attr.IsPolymorphic ? ' (polymorphic)' : ''} | ${tcell(bind)} |`
    );
  });
  lines.push('');
  return lines;
}

function buildOptionSetSection(
  fields: AttributeMetadataComplete[],
  verbosity: VerbosityLevel
): string[] {
  const optionFields = fields.filter(f => f.OptionSetValues && f.OptionSetValues.length > 0);
  if (optionFields.length === 0) return [];

  const lines: string[] = [];
  lines.push('### Option Sets');
  lines.push('');

  const optionLimit = verbosity === 'compact' ? 0 : verbosity === 'standard' ? 25 : Infinity;
  if (verbosity === 'compact') {
    lines.push('_(option-set values omitted in compact verbosity — switch to standard / full to include them)_');
    lines.push('');
    return lines;
  }

  optionFields.forEach(attr => {
    lines.push(`#### \`${attr.LogicalName}\` — ${attr.DisplayName} (${attr.AttributeType})`);
    const opts = attr.OptionSetValues || [];
    if (opts.length <= optionLimit) {
      opts.forEach(opt => lines.push(`- \`${opt.Value}\` = ${opt.Label}`));
    } else {
      opts.slice(0, optionLimit).forEach(opt => lines.push(`- \`${opt.Value}\` = ${opt.Label}`));
      lines.push(`- _(${opts.length - optionLimit} more values truncated — use full verbosity to include all)_`);
    }
    lines.push('');
  });
  return lines;
}

function buildExampleQueries(
  entityMeta: EntityMetadataComplete,
  selectedFields: AttributeMetadataComplete[],
  apiVersion: string,
  baseUrl: string | null,
  verbosity: VerbosityLevel
): string[] {
  if (verbosity === 'compact') return [];

  const lines: string[] = [];
  const entitySet = entityMeta.EntitySetName;
  const primaryId = entityMeta.PrimaryIdAttribute;
  const primaryName = entityMeta.PrimaryNameAttribute;
  const root = baseUrl ? `${baseUrl}/api/data/${apiVersion}` : `/api/data/${apiVersion}`;
  const apiPath = `${root}/${entitySet}`;

  // Build a $select that uses the lookup _value form for any lookup columns.
  const selectParts = selectedFields
    .filter(f => f.IsValidForRead !== false && !f.AttributeOf) // skip read-blocked + shadow attrs
    .slice(0, 6) // keep example queries short
    .map(f => LOOKUP_TYPES.has(f.AttributeType) ? `_${f.LogicalName}_value` : f.LogicalName);
  // Always include PK + primary name if they exist
  if (primaryId && !selectParts.includes(primaryId)) selectParts.unshift(primaryId);
  if (primaryName && !selectParts.includes(primaryName)) selectParts.splice(1, 0, primaryName);
  const selectStr = selectParts.length > 0 ? selectParts.join(',') : primaryId;

  // Pick a sample lookup, option set, and date for filter examples.
  const sampleLookup = selectedFields.find(f => LOOKUP_TYPES.has(f.AttributeType) && !f.AttributeOf);
  const sampleOption = selectedFields.find(f => f.OptionSetValues && f.OptionSetValues.length > 0);
  const sampleDate = selectedFields.find(f => f.AttributeType === 'DateTime');

  lines.push('### Example Queries');
  lines.push('');

  lines.push('Top 10 records:');
  lines.push('```http');
  lines.push(`GET ${apiPath}?$select=${selectStr}&$top=10`);
  lines.push('```');
  lines.push('');

  lines.push('Retrieve by primary id:');
  lines.push('```http');
  lines.push(`GET ${apiPath}(<guid>)?$select=${selectStr}`);
  lines.push('```');
  lines.push('');

  if (sampleLookup) {
    lines.push(`Filter by lookup (${sampleLookup.LogicalName}):`);
    lines.push('```http');
    lines.push(`GET ${apiPath}?$select=${selectStr}&$filter=_${sampleLookup.LogicalName}_value eq <guid>`);
    lines.push('```');
    lines.push('');
  }

  if (sampleDate) {
    lines.push(`Filter by date (${sampleDate.LogicalName}):`);
    lines.push('```http');
    lines.push(`GET ${apiPath}?$select=${selectStr}&$filter=${sampleDate.LogicalName} ge 2025-01-01T00:00:00Z`);
    lines.push('```');
    lines.push('');
  }

  if (sampleOption && sampleOption.OptionSetValues && sampleOption.OptionSetValues.length > 0) {
    const sampleVal = sampleOption.OptionSetValues[0].Value;
    const sampleLbl = sampleOption.OptionSetValues[0].Label;
    lines.push(`Filter by option set (${sampleOption.LogicalName} = ${sampleLbl}):`);
    lines.push('```http');
    lines.push(`GET ${apiPath}?$select=${selectStr}&$filter=${sampleOption.LogicalName} eq ${sampleVal}`);
    lines.push('```');
    lines.push('');
  }

  if (verbosity === 'full') {
    // Create example with first writable required-or-recommended field + first lookup bind
    const writable = selectedFields.filter(f =>
      f.IsValidForCreate !== false && !f.AttributeOf && !f.IsPrimaryId &&
      f.AttributeType !== 'Uniqueidentifier'
    );
    const firstScalar = writable.find(f => !LOOKUP_TYPES.has(f.AttributeType));
    const firstLookup = writable.find(f => LOOKUP_TYPES.has(f.AttributeType) && f.LookupTargets && f.LookupTargets.length > 0);

    if (firstScalar || firstLookup) {
      lines.push('Create:');
      lines.push('```http');
      lines.push(`POST ${apiPath}`);
      lines.push('Content-Type: application/json');
      lines.push('Prefer: return=representation');
      lines.push('');
      const body: string[] = [];
      if (firstScalar) body.push(`  "${firstScalar.LogicalName}": ${exampleValue(firstScalar)}`);
      if (firstLookup && firstLookup.LookupTargets && firstLookup.LookupTargets[0]) {
        const target = firstLookup.LookupTargets[0];
        body.push(`  "${firstLookup.LogicalName}@odata.bind": "/${target}s(<guid>)"`);
      }
      lines.push('{');
      lines.push(body.join(',\n'));
      lines.push('}');
      lines.push('```');
      lines.push('');
    }

    lines.push('Update single field:');
    lines.push('```http');
    lines.push(`PATCH ${apiPath}(<guid>)`);
    lines.push('Content-Type: application/json');
    lines.push('');
    if (firstScalar) {
      lines.push(`{ "${firstScalar.LogicalName}": ${exampleValue(firstScalar)} }`);
    } else {
      lines.push(`{ "${primaryName || primaryId}": "..." }`);
    }
    lines.push('```');
    lines.push('');

    lines.push('Delete:');
    lines.push('```http');
    lines.push(`DELETE ${apiPath}(<guid>)`);
    lines.push('```');
    lines.push('');
  }

  if (entityMeta.LogicalName === 'activitypointer') {
    lines.push('Activities for a parent record (regarding):');
    lines.push('```http');
    lines.push(`GET ${root}/activitypointers?$select=activityid,subject,_regardingobjectid_value&$filter=_regardingobjectid_value eq <guid>&$top=10`);
    lines.push('```');
    lines.push('');
  }

  return lines;
}

function exampleValue(attr: AttributeMetadataComplete): string {
  switch (attr.AttributeType) {
    case 'String':
    case 'Memo':
      return '"example"';
    case 'Integer':
    case 'BigInt':
      return '0';
    case 'Decimal':
    case 'Double':
    case 'Money':
      return '0.0';
    case 'Boolean':
      return 'false';
    case 'DateTime':
      return '"2025-01-01T00:00:00Z"';
    case 'Picklist':
    case 'State':
    case 'Status':
      return attr.OptionSetValues && attr.OptionSetValues[0] ? String(attr.OptionSetValues[0].Value) : '0';
    default:
      return '"..."';
  }
}

/** Build a quick set of target -> entitySet lookups across all selected entities so
 *  lookup bind examples can use the right collection name. */
function buildTargetEntitySetMap(
  metadataMap: Map<string, EntityMetadataComplete>
): Map<string, string> {
  const map = new Map<string, string>();
  metadataMap.forEach(meta => {
    if (meta.LogicalName && meta.EntitySetName) {
      map.set(meta.LogicalName, meta.EntitySetName);
    }
  });
  return map;
}

export function generatePromptMarkdown(
  metadataMap: Map<string, EntityMetadataComplete>,
  selections: PromptSelections,
  options: PromptOptions
): string {
  const lines: string[] = [];
  const baseUrl = getOrgBaseUrl(options.orgUrl);
  const apiVersion = getApiVersion(options);
  const targetEntitySetMap = buildTargetEntitySetMap(metadataMap);

  // Header
  if (options.includeRules) {
    lines.push('# Dynamics 365 Entity Context');
    lines.push('');
    lines.push('## Context Information');
    if (options.orgUrl) {
      try {
        lines.push(`- **Environment**: ${new URL(options.orgUrl).hostname}`);
      } catch {
        // ignore malformed orgUrl
      }
    }
    if (baseUrl) {
      lines.push(`- **Web API Base**: \`${baseUrl}/api/data/${apiVersion}\``);
    } else {
      lines.push(`- **Web API Base**: \`/api/data/${apiVersion}\``);
    }
    if (options.generatedAt) {
      lines.push(`- **Generated**: ${options.generatedAt}`);
    }
    lines.push('');

    // Cheat sheet — the highest-value section for an LLM consuming this prompt
    lines.push(...buildCheatSheet());

    // Quick reference table of all selected entities
    lines.push('## Selected Entities');
    lines.push('');
    lines.push('| Display | Logical | Entity Set | Primary ID | Primary Name |');
    lines.push('|---|---|---|---|---|');
    selections.entities.forEach((_selection, logicalName) => {
      const meta = metadataMap.get(logicalName);
      if (meta) {
        lines.push(
          `| ${tcell(meta.DisplayName)} | \`${tcell(meta.LogicalName)}\` | \`${tcell(meta.EntitySetName)}\` | \`${tcell(meta.PrimaryIdAttribute)}\` | ${meta.PrimaryNameAttribute ? `\`${tcell(meta.PrimaryNameAttribute)}\`` : '_(none)_'} |`
        );
      } else {
        lines.push(`| _(metadata not loaded)_ | \`${tcell(logicalName)}\` | | | |`);
      }
    });
    lines.push('');
  }

  // Per-entity sections
  selections.entities.forEach((selection, entityLogicalName) => {
    const entityMeta = metadataMap.get(entityLogicalName);
    if (!entityMeta) {
      lines.push(`## ${entityLogicalName} Entity Details`);
      lines.push('');
      lines.push(`> _Metadata for \`${entityLogicalName}\` was not loaded; details unavailable._`);
      lines.push('');
      return;
    }

    lines.push(`## ${entityMeta.DisplayName} Entity Details`);
    lines.push('');
    lines.push(`- **Logical Name**: \`${entityMeta.LogicalName}\``);
    lines.push(`- **Entity Set**: \`${entityMeta.EntitySetName}\``);
    lines.push(`- **Primary ID**: \`${entityMeta.PrimaryIdAttribute}\``);
    if (entityMeta.PrimaryNameAttribute) {
      lines.push(`- **Primary Name**: \`${entityMeta.PrimaryNameAttribute}\``);
    }
    if (entityMeta.Description) {
      lines.push(`- **Description**: ${entityMeta.Description}`);
    }
    lines.push('');

    const selectedSet = new Set(selection.selectedFields);
    const selectedFields = entityMeta.Attributes
      .filter(attr => selectedSet.has(attr.LogicalName))
      .sort((a, b) => {
        // PK first, then primary name, then alphabetical
        if (a.IsPrimaryId !== b.IsPrimaryId) return a.IsPrimaryId ? -1 : 1;
        if (a.IsPrimaryName !== b.IsPrimaryName) return a.IsPrimaryName ? -1 : 1;
        return (a.LogicalName || '').localeCompare(b.LogicalName || '');
      });

    if (selectedFields.length > 0) {
      lines.push('### Field Reference');
      lines.push('');
      lines.push(...buildFieldTable(selectedFields, options.verbosity));
      lines.push(...buildLookupTable(selectedFields, targetEntitySetMap));
      lines.push(...buildOptionSetSection(selectedFields, options.verbosity));
    } else {
      lines.push('_No fields selected for this entity._');
      lines.push('');
    }

    lines.push(...buildExampleQueries(entityMeta, selectedFields, apiVersion, baseUrl, options.verbosity));
  });

  // Relationships — collect each globally-selected relationship exactly once.
  // Same SchemaName can appear under both OneToMany (parent) and ManyToOne (child);
  // canonicalize to OneToMany so direction/type are stable.
  const allRelationships: RelationshipMetadata[] = [];
  const selectedRelSet = new Set(selections.selectedRelationships);
  if (selectedRelSet.size > 0) {
    const candidates = new Map<string, RelationshipMetadata>();
    const rank = (t: RelationshipMetadata['RelationshipType']) =>
      t === 'OneToMany' ? 0 : t === 'ManyToMany' ? 1 : 2;

    metadataMap.forEach(entityMeta => {
      [
        ...entityMeta.OneToManyRelationships,
        ...entityMeta.ManyToManyRelationships,
        ...entityMeta.ManyToOneRelationships
      ].forEach(rel => {
        if (!selectedRelSet.has(rel.SchemaName)) return;
        const existing = candidates.get(rel.SchemaName);
        if (!existing || rank(rel.RelationshipType) < rank(existing.RelationshipType)) {
          candidates.set(rel.SchemaName, rel);
        }
      });
    });
    candidates.forEach(rel => allRelationships.push(rel));
  }

  if (allRelationships.length > 0) {
    lines.push('## Relationships');
    lines.push('');

    const byType: Record<string, RelationshipMetadata[]> = {
      OneToMany: [],
      ManyToOne: [],
      ManyToMany: []
    };
    allRelationships.forEach(rel => byType[rel.RelationshipType].push(rel));

    Object.entries(byType).forEach(([type, rels]) => {
      if (rels.length === 0) return;
      lines.push(`### ${type} Relationships`);
      lines.push('');

      // Table for compact-friendly listing of the essentials
      lines.push('| Schema | Direction | Lookup Attribute | Nav Property (for $expand) | $expand example |');
      lines.push('|---|---|---|---|---|');

      rels.forEach(rel => {
        const childMeta = metadataMap.get(rel.ReferencingEntity);
        const parentMeta = metadataMap.get(rel.ReferencedEntity);
        const childName = childMeta?.DisplayName || rel.ReferencingEntity;
        const parentName = parentMeta?.DisplayName || rel.ReferencedEntity;
        const childSet = childMeta?.EntitySetName || rel.ReferencingEntity;
        const parentSet = parentMeta?.EntitySetName || rel.ReferencedEntity;
        const isSelf = rel.ReferencingEntity === rel.ReferencedEntity;

        let direction: string;
        let nav: string | undefined;
        let example: string;
        let lookupAttr: string;

        if (isSelf) {
          direction = `${parentName} (self)`;
        } else if (type === 'OneToMany') {
          direction = `${parentName} → ${childName}`;
        } else if (type === 'ManyToOne') {
          direction = `${childName} → ${parentName}`;
        } else {
          direction = `${parentName} ↔ ${childName}`;
        }

        if (type === 'OneToMany') {
          nav = rel.ReferencedEntityNavigationPropertyName; // collection on parent
          lookupAttr = rel.ReferencingAttribute || '';
          example = nav ? `/${parentSet}?$expand=${nav}($select=...)` : '_(no nav property)_';
        } else if (type === 'ManyToOne') {
          nav = rel.ReferencingEntityNavigationPropertyName; // single-valued on child
          lookupAttr = rel.ReferencingAttribute || '';
          example = nav ? `/${childSet}?$expand=${nav}($select=...)` : '_(no nav property)_';
        } else {
          // ManyToMany — show both nav props if available
          const navs: string[] = [];
          if (rel.ReferencingEntityNavigationPropertyName) {
            navs.push(`on ${childName}: \`${rel.ReferencingEntityNavigationPropertyName}\``);
          }
          if (rel.ReferencedEntityNavigationPropertyName) {
            navs.push(`on ${parentName}: \`${rel.ReferencedEntityNavigationPropertyName}\``);
          }
          nav = navs.join(' / ');
          lookupAttr = '';
          const exa = rel.ReferencingEntityNavigationPropertyName
            ? `/${childSet}?$expand=${rel.ReferencingEntityNavigationPropertyName}($select=...)`
            : '_(no nav property)_';
          example = exa;
        }

        lines.push(
          `| \`${tcell(rel.SchemaName)}\` | ${tcell(direction)} | ${lookupAttr ? `\`${tcell(lookupAttr)}\`` : ''} | ${nav ? (nav.startsWith('on ') ? tcell(nav) : `\`${tcell(nav)}\``) : ''} | \`${tcell(example)}\` |`
        );
      });
      lines.push('');
    });
  }

  return lines.join('\n');
}
