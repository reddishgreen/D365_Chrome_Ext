import { SelectedEntity, JoinedEntity, QueryColumn, QueryFilter } from '../types';

export const buildODataQuery = (
  mainEntity: SelectedEntity,
  joinedEntities: JoinedEntity[],
  columns: QueryColumn[],
  filters: QueryFilter
): string => {
  // 1. Build Select and Expand
  // We need to construct a tree of expands.
  // main -> [child1, child2]
  // child1 -> [grandchild1]
  
  const buildExpandTree = (parentAlias: string): string => {
    const children = joinedEntities.filter(j => j.parentAlias === parentAlias);
    if (children.length === 0) return '';

    const expands = children.map(child => {
      // Get columns for this child
      const childCols = columns.filter(c => c.entityAlias === child.alias).map(c => c.attribute);
      // Always include primary ID for linking
      if (!childCols.includes(child.primaryIdAttribute)) {
        childCols.push(child.primaryIdAttribute);
      }
      
      let select = childCols.length > 0 ? `$select=${childCols.join(',')}` : '';
      
      // Recursively build nested expands
      const nestedExpand = buildExpandTree(child.alias);
      let expand = nestedExpand ? `$expand=${nestedExpand}` : '';

      // Combine select and expand
      let options = [];
      if (select) options.push(select);
      if (expand) options.push(expand);
      
      const optionsStr = options.length > 0 ? `(${options.join(';')})` : '';
      
      return `${child.navigationPropertyName}${optionsStr}`;
    });

    return expands.join(',');
  };

  const expandStr = buildExpandTree(mainEntity.alias);

  // Main entity columns
  const mainCols = columns.filter(c => c.entityAlias === mainEntity.alias).map(c => c.attribute);
  if (!mainCols.includes(mainEntity.primaryIdAttribute)) {
    mainCols.push(mainEntity.primaryIdAttribute);
  }
  const selectStr = `$select=${mainCols.join(',')}`;

  // 2. Build Filter
  const buildFilterString = (node: QueryFilter): string => {
    if (node.type === 'group') {
      if (!node.children || node.children.length === 0) return '';
      
      const childFilters = node.children
        .map(buildFilterString)
        .filter(f => f !== '');
      
      if (childFilters.length === 0) return '';
      if (childFilters.length === 1) return childFilters[0];
      
      return `(${childFilters.join(` ${node.logicalOperator} `)})`;
    } else {
      // Condition
      if (!node.attribute || !node.operator) return ''; // Incomplete

      // Resolve path to attribute
      let attributePath = node.attribute;
      
      if (node.entityAlias && node.entityAlias !== 'main') {
         // Find path from main to this alias
         // For now, we assume simple paths or we need to look up the hierarchy
         // joinedEntities has parentAlias.
         // We need to walk up from node.entityAlias to 'main'
         const path: string[] = [];
         let currentAlias = node.entityAlias;
         
         while (currentAlias !== 'main') {
           const entity = joinedEntities.find(j => j.alias === currentAlias);
           if (!entity) break; // Should not happen
           path.unshift(entity.navigationPropertyName);
           currentAlias = entity.parentAlias;
         }
         
         if (path.length > 0) {
           attributePath = `${path.join('/')}/${node.attribute}`;
         }
      }

      // Handle operators
      const val = node.value;
      
      // Handle string quoting
      const formatValue = (v: any) => {
         if (typeof v === 'string') return `'${v}'`;
         return v;
      };

      switch (node.operator) {
        // Basic Operators
        case 'eq': return `${attributePath} eq ${formatValue(val)}`;
        case 'ne': return `${attributePath} ne ${formatValue(val)}`;
        case 'null': return `${attributePath} eq null`;
        case 'not null': return `${attributePath} ne null`;

        // String Operators
        case 'contains': return `contains(${attributePath}, ${formatValue(val)})`;
        case 'not contains': return `not (contains(${attributePath}, ${formatValue(val)}))`;
        case 'startswith': return `startswith(${attributePath}, ${formatValue(val)})`;
        case 'endswith': return `endswith(${attributePath}, ${formatValue(val)})`;

        // Numeric & Date Comparison Operators
        case 'gt': return `${attributePath} gt ${formatValue(val)}`;
        case 'ge': return `${attributePath} ge ${formatValue(val)}`;
        case 'lt': return `${attributePath} lt ${formatValue(val)}`;
        case 'le': return `${attributePath} le ${formatValue(val)}`;

        // Date-Specific Operators
        case 'on': return `Microsoft.Dynamics.CRM.On(PropertyName='${attributePath}',PropertyValue='${val}')`;
        case 'on-or-after': return `Microsoft.Dynamics.CRM.OnOrAfter(PropertyName='${attributePath}',PropertyValue='${val}')`;
        case 'on-or-before': return `Microsoft.Dynamics.CRM.OnOrBefore(PropertyName='${attributePath}',PropertyValue='${val}')`;

        // Relative Date Operators - Daily
        case 'today': return `Microsoft.Dynamics.CRM.Today(PropertyName='${attributePath}')`;
        case 'yesterday': return `Microsoft.Dynamics.CRM.Yesterday(PropertyName='${attributePath}')`;
        case 'tomorrow': return `Microsoft.Dynamics.CRM.Tomorrow(PropertyName='${attributePath}')`;
        case 'last-x-days': return `Microsoft.Dynamics.CRM.LastXDays(PropertyName='${attributePath}',PropertyValue=${val})`;
        case 'next-x-days': return `Microsoft.Dynamics.CRM.NextXDays(PropertyName='${attributePath}',PropertyValue=${val})`;
        case 'older-than-x-days': return `Microsoft.Dynamics.CRM.OlderThanXDays(PropertyName='${attributePath}',PropertyValue=${val})`;

        // Relative Date Operators - Weekly
        case 'this-week': return `Microsoft.Dynamics.CRM.ThisWeek(PropertyName='${attributePath}')`;
        case 'last-week': return `Microsoft.Dynamics.CRM.LastWeek(PropertyName='${attributePath}')`;
        case 'next-week': return `Microsoft.Dynamics.CRM.NextWeek(PropertyName='${attributePath}')`;
        case 'last-x-weeks': return `Microsoft.Dynamics.CRM.LastXWeeks(PropertyName='${attributePath}',PropertyValue=${val})`;
        case 'next-x-weeks': return `Microsoft.Dynamics.CRM.NextXWeeks(PropertyName='${attributePath}',PropertyValue=${val})`;

        // Relative Date Operators - Monthly
        case 'this-month': return `Microsoft.Dynamics.CRM.ThisMonth(PropertyName='${attributePath}')`;
        case 'last-month': return `Microsoft.Dynamics.CRM.LastMonth(PropertyName='${attributePath}')`;
        case 'next-month': return `Microsoft.Dynamics.CRM.NextMonth(PropertyName='${attributePath}')`;
        case 'last-x-months': return `Microsoft.Dynamics.CRM.LastXMonths(PropertyName='${attributePath}',PropertyValue=${val})`;
        case 'next-x-months': return `Microsoft.Dynamics.CRM.NextXMonths(PropertyName='${attributePath}',PropertyValue=${val})`;
        case 'older-than-x-months': return `Microsoft.Dynamics.CRM.OlderThanXMonths(PropertyName='${attributePath}',PropertyValue=${val})`;

        // Relative Date Operators - Yearly
        case 'this-year': return `Microsoft.Dynamics.CRM.ThisYear(PropertyName='${attributePath}')`;
        case 'last-year': return `Microsoft.Dynamics.CRM.LastYear(PropertyName='${attributePath}')`;
        case 'next-year': return `Microsoft.Dynamics.CRM.NextYear(PropertyName='${attributePath}')`;
        case 'last-x-years': return `Microsoft.Dynamics.CRM.LastXYears(PropertyName='${attributePath}',PropertyValue=${val})`;
        case 'next-x-years': return `Microsoft.Dynamics.CRM.NextXYears(PropertyName='${attributePath}',PropertyValue=${val})`;
        case 'older-than-x-years': return `Microsoft.Dynamics.CRM.OlderThanXYears(PropertyName='${attributePath}',PropertyValue=${val})`;

        // Fiscal Period Operators
        case 'this-fiscal-year': return `Microsoft.Dynamics.CRM.InFiscalYear(PropertyName='${attributePath}',PropertyValue=0)`;
        case 'this-fiscal-period': return `Microsoft.Dynamics.CRM.InFiscalPeriod(PropertyName='${attributePath}',PropertyValue=0)`;
        case 'last-fiscal-year': return `Microsoft.Dynamics.CRM.InFiscalYear(PropertyName='${attributePath}',PropertyValue=-1)`;
        case 'last-fiscal-period': return `Microsoft.Dynamics.CRM.InFiscalPeriod(PropertyName='${attributePath}',PropertyValue=-1)`;
        case 'next-fiscal-year': return `Microsoft.Dynamics.CRM.InFiscalYear(PropertyName='${attributePath}',PropertyValue=1)`;
        case 'next-fiscal-period': return `Microsoft.Dynamics.CRM.InFiscalPeriod(PropertyName='${attributePath}',PropertyValue=1)`;

        default: return '';
      }
    }
  };

  const filterStr = buildFilterString(filters);
  
  // Combine all - build query string
  // Note: OData query parameters should be properly formatted
  // Commas in $select should NOT be encoded, but the fetch API will handle encoding
  let query = `${mainEntity.entitySetName}?${selectStr}`;
  if (expandStr) {
    query += `&$expand=${expandStr}`;
  }
  if (filterStr) {
    query += `&$filter=${encodeURIComponent(filterStr)}`;
  }
  
  return query;
};

