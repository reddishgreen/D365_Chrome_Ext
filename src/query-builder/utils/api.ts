import { EntityMetadata, AttributeMetadata, RelationshipMetadata, ViewMetadata } from '../types';

export class CrmApi {
  private baseUrl: string;

  constructor(orgUrl: string) {
    this.baseUrl = `${orgUrl}/api/data/v9.2`;
  }

  private async fetchJson(url: string): Promise<any> {
    const response = await fetch(url, {
      headers: {
        'Accept': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        'Prefer': 'odata.include-annotations="*"'
      },
      credentials: 'include'
    });

    if (!response.ok) {
      let errorDetails = response.statusText;
      try {
        const text = await response.text();
        if (text) {
          try {
            const json = JSON.parse(text);
            errorDetails = json.error?.message || json.Message || text;
          } catch {
            errorDetails = text;
          }
        }
      } catch {
        // Ignore error reading text
      }
      throw new Error(`CRM API Error (${response.status}): ${errorDetails}`);
    }

    return await response.json();
  }

  async getAllEntities(): Promise<EntityMetadata[]> {
    // Fetching minimal properties to avoid 400 errors with unsupported select/orderby
    const url = `${this.baseUrl}/EntityDefinitions?$select=LogicalName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute`;
    
    try {
      const data = await this.fetchJson(url);
      const entities = data.value.map((e: any) => ({
        LogicalName: e.LogicalName,
        DisplayName: e.DisplayName?.UserLocalizedLabel?.Label || e.LogicalName,
        EntitySetName: e.EntitySetName,
        PrimaryIdAttribute: e.PrimaryIdAttribute,
        PrimaryNameAttribute: e.PrimaryNameAttribute
      }));

      return entities.sort((a: EntityMetadata, b: EntityMetadata) => 
        (a.DisplayName || a.LogicalName).localeCompare(b.DisplayName || b.LogicalName)
      );
    } catch (e) {
      // Fallback: Try without $select if specific properties are causing issues
      console.warn('Failed with $select, retrying without query parameters', e);
      const fallbackUrl = `${this.baseUrl}/EntityDefinitions`;
      const data = await this.fetchJson(fallbackUrl);
      
      const entities = data.value.map((e: any) => ({
        LogicalName: e.LogicalName,
        DisplayName: e.DisplayName?.UserLocalizedLabel?.Label || e.LogicalName,
        EntitySetName: e.EntitySetName,
        PrimaryIdAttribute: e.PrimaryIdAttribute,
        PrimaryNameAttribute: e.PrimaryNameAttribute
      }));

      return entities.sort((a: EntityMetadata, b: EntityMetadata) => 
        (a.DisplayName || a.LogicalName).localeCompare(b.DisplayName || b.LogicalName)
      );
    }
  }

  async getEntity(logicalName: string): Promise<EntityMetadata> {
    const url = `${this.baseUrl}/EntityDefinitions(LogicalName='${logicalName}')?$select=LogicalName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute`;
    
    const e = await this.fetchJson(url);
    return {
      LogicalName: e.LogicalName,
      DisplayName: e.DisplayName?.UserLocalizedLabel?.Label || e.LogicalName,
      EntitySetName: e.EntitySetName,
      PrimaryIdAttribute: e.PrimaryIdAttribute,
      PrimaryNameAttribute: e.PrimaryNameAttribute
    };
  }

  async getAttributes(logicalName: string): Promise<AttributeMetadata[]> {
    // $orderby is not supported on Attributes in some versions
    const url = `${this.baseUrl}/EntityDefinitions(LogicalName='${logicalName}')/Attributes?$select=LogicalName,DisplayName,AttributeType,IsPrimaryId,IsPrimaryName`;
    const data = await this.fetchJson(url);
    
    const attributes = data.value.map((a: any) => ({
      LogicalName: a.LogicalName,
      DisplayName: a.DisplayName?.UserLocalizedLabel?.Label || a.LogicalName,
      AttributeType: a.AttributeType,
      IsPrimaryId: a.IsPrimaryId,
      IsPrimaryName: a.IsPrimaryName
    }));

    return attributes.sort((a: AttributeMetadata, b: AttributeMetadata) => 
      (a.DisplayName || a.LogicalName).localeCompare(b.DisplayName || b.LogicalName)
    );
  }

  async getRelationships(logicalName: string, type: 'OneToMany' | 'ManyToOne'): Promise<RelationshipMetadata[]> {
    const url = `${this.baseUrl}/EntityDefinitions(LogicalName='${logicalName}')/${type}Relationships`;
    const data = await this.fetchJson(url);

    return data.value.map((r: any) => ({
      SchemaName: r.SchemaName,
      ReferencingEntity: r.ReferencingEntity,
      ReferencedEntity: r.ReferencedEntity,
      ReferencingAttribute: r.ReferencingAttribute,
      // For ManyToOne (Lookup), the nav prop is on the Referencing Entity (us).
      // For OneToMany, the nav prop is on the Referenced Entity (us).
      ReferencingEntityNavigationPropertyName: r.ReferencingEntityNavigationPropertyName, // On Child
      ReferencedEntityNavigationPropertyName: r.ReferencedEntityNavigationPropertyName, // On Parent
      RelationshipType: type
    }));
  }

  async getViews(entityLogicalName: string): Promise<ViewMetadata[]> {
    // 1. Fetch System Views (savedquery)
    // querytype=0 (Main Application View)
    const systemViewsUrl = `${this.baseUrl}/savedqueries?$select=name,savedqueryid,fetchxml,querytype&$filter=returnedtypecode eq '${entityLogicalName}' and querytype eq 0`;
    
    // 2. Fetch Personal Views (userquery)
    const userViewsUrl = `${this.baseUrl}/userqueries?$select=name,userqueryid,fetchxml,querytype&$filter=returnedtypecode eq '${entityLogicalName}'`;

    const [systemData, userData] = await Promise.all([
      this.fetchJson(systemViewsUrl).catch(() => ({ value: [] })),
      this.fetchJson(userViewsUrl).catch(() => ({ value: [] }))
    ]);

    const systemViews: ViewMetadata[] = (systemData.value || []).map((v: any) => ({
      id: v.savedqueryid,
      name: v.name,
      fetchXml: v.fetchxml,
      queryType: v.querytype,
      isUserQuery: false
    }));

    const userViews: ViewMetadata[] = (userData.value || []).map((v: any) => ({
      id: v.userqueryid,
      name: v.name,
      fetchXml: v.fetchxml,
      queryType: v.querytype,
      isUserQuery: true
    }));

    // Combine and sort by name
    return [...systemViews, ...userViews].sort((a, b) => a.name.localeCompare(b.name));
  }

  async executeQuery(queryUrl: string): Promise<any> {
    let url = queryUrl;
    if (!url.startsWith('http')) {
        url = `${this.baseUrl}/${url}`;
    }
    return await this.fetchJson(url);
  }
}
