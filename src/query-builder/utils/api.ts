import { EntityMetadata, AttributeMetadata, RelationshipMetadata, ViewMetadata, EntityMetadataComplete, AttributeMetadataComplete } from '../types';

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

  async getOptionSetValues(entityLogicalName: string, attributeLogicalName: string): Promise<Array<{ Value: number; Label: string }>> {
    // Try each attribute type in sequence until one works
    const tryFetchOptions = async (metadataType: string): Promise<Array<{ Value: number; Label: string }> | null> => {
      try {
        const url = `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attributeLogicalName}')/${metadataType}?$expand=OptionSet`;
        const data = await this.fetchJson(url);

        if (data.OptionSet?.Options) {
          // Standard option set (Picklist, State, Status)
          return data.OptionSet.Options.map((opt: any) => ({
            Value: opt.Value,
            Label: opt.Label?.UserLocalizedLabel?.Label || String(opt.Value)
          }));
        } else if (data.OptionSet?.FalseOption || data.OptionSet?.TrueOption) {
          // Boolean attribute - has FalseOption/TrueOption instead of Options array
          const options: Array<{ Value: number; Label: string }> = [];
          if (data.OptionSet.FalseOption) {
            options.push({
              Value: data.OptionSet.FalseOption.Value,
              Label: data.OptionSet.FalseOption.Label?.UserLocalizedLabel?.Label || 'No'
            });
          }
          if (data.OptionSet.TrueOption) {
            options.push({
              Value: data.OptionSet.TrueOption.Value,
              Label: data.OptionSet.TrueOption.Label?.UserLocalizedLabel?.Label || 'Yes'
            });
          }
          return options;
        }
        return null;
      } catch {
        return null;
      }
    };

    // Try each metadata type in order
    const metadataTypes = [
      'Microsoft.Dynamics.CRM.PicklistAttributeMetadata',
      'Microsoft.Dynamics.CRM.StateAttributeMetadata',
      'Microsoft.Dynamics.CRM.StatusAttributeMetadata',
      'Microsoft.Dynamics.CRM.BooleanAttributeMetadata'
    ];

    for (const metadataType of metadataTypes) {
      const options = await tryFetchOptions(metadataType);
      if (options && options.length > 0) {
        return options;
      }
    }

    return [];
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

  /**
   * Fetches all pages of data from an OData query, handling @odata.nextLink pagination.
   * This bypasses the default 5000 record limit by following continuation links.
   *
   * @param queryUrl - The initial OData query URL
   * @param onProgress - Optional callback to report progress (recordsFetched, isComplete)
   * @returns Promise resolving to all records concatenated together
   */
  async fetchAllPages(
    queryUrl: string,
    onProgress?: (recordsFetched: number, isComplete: boolean) => void
  ): Promise<any[]> {
    let allRecords: any[] = [];
    let url = queryUrl;

    if (!url.startsWith('http')) {
      url = `${this.baseUrl}/${url}`;
    }

    while (url) {
      const response = await this.fetchJson(url);
      const records = response.value || [];
      allRecords = allRecords.concat(records);

      // Report progress
      if (onProgress) {
        onProgress(allRecords.length, !response['@odata.nextLink']);
      }

      // Check for next page
      url = response['@odata.nextLink'] || null;
    }

    return allRecords;
  }

  async getEntityMetadata(entityLogicalName: string): Promise<EntityMetadataComplete> {
    // Fetch entity definition
    const entityUrl = `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')?$select=LogicalName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,Description`;
    const entityData = await this.fetchJson(entityUrl);

    // Fetch all attributes with comprehensive metadata
    // Note: MaxLength, Precision, Scale may not be available on all attribute types
    // We'll fetch them separately for specific attribute types that support them
    const attributesUrl = `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes?$select=LogicalName,DisplayName,AttributeType,IsPrimaryId,IsPrimaryName,RequiredLevel,Description`;
    const attributesData = await this.fetchJson(attributesUrl);

    // Process attributes and fetch additional details for option sets and lookups
    const attributes: AttributeMetadataComplete[] = await Promise.all(
      attributesData.value.map(async (attr: any) => {
        const attribute: AttributeMetadataComplete = {
          LogicalName: attr.LogicalName,
          DisplayName: attr.DisplayName?.UserLocalizedLabel?.Label || attr.LogicalName,
          AttributeType: attr.AttributeType,
          IsPrimaryId: attr.IsPrimaryId,
          IsPrimaryName: attr.IsPrimaryName,
          RequiredLevel: attr.RequiredLevel?.Value,
          Description: attr.Description?.UserLocalizedLabel?.Label,
          IsCalculated: attr.IsCalculated,
          IsRollup: attr.IsRollup
        };

        // Fetch MaxLength, Precision, Scale for specific attribute types that support them
        if (attr.AttributeType === 'String' || attr.AttributeType === 'Memo') {
          try {
            const stringUrl = `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attr.LogicalName}')/Microsoft.Dynamics.CRM.StringAttributeMetadata?$select=MaxLength`;
            const stringData = await this.fetchJson(stringUrl);
            attribute.MaxLength = stringData.MaxLength;
          } catch (e) {
            // MaxLength not available for this attribute type, skip it
          }
        } else if (attr.AttributeType === 'Decimal' || attr.AttributeType === 'Money') {
          try {
            const decimalUrl = `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attr.LogicalName}')/Microsoft.Dynamics.CRM.DecimalAttributeMetadata?$select=Precision`;
            const decimalData = await this.fetchJson(decimalUrl);
            attribute.Precision = decimalData.Precision;
          } catch (e) {
            // Precision not available, skip it
          }
        }

        // Fetch option set values if it's an option set
        if (attr.AttributeType === 'Picklist' || attr.AttributeType === 'State' || attr.AttributeType === 'Status') {
          try {
            const optionSetUrl = `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attr.LogicalName}')/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$expand=OptionSet($expand=Options($select=Value,Label))`;
            const optionSetData = await this.fetchJson(optionSetUrl);
            if (optionSetData.OptionSet?.Options) {
              attribute.OptionSetValues = optionSetData.OptionSet.Options.map((opt: any) => ({
                Value: opt.Value,
                Label: opt.Label?.UserLocalizedLabel?.Label || String(opt.Value)
              }));
            }
          } catch (e) {
            // Option set fetch failed, continue without values
            console.warn(`Failed to fetch option set for ${attr.LogicalName}`, e);
          }
        }

        // Fetch lookup targets if it's a lookup
        if (attr.AttributeType === 'Lookup') {
          try {
            const lookupUrl = `${this.baseUrl}/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes(LogicalName='${attr.LogicalName}')/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=Targets`;
            const lookupData = await this.fetchJson(lookupUrl);
            attribute.LookupTargets = lookupData.Targets || [];
            
            // Check if polymorphic (Customer/Regarding)
            if (lookupData.Targets && lookupData.Targets.length > 1) {
              attribute.IsPolymorphic = true;
            }
          } catch (e) {
            // Lookup fetch failed, continue without targets
            console.warn(`Failed to fetch lookup targets for ${attr.LogicalName}`, e);
          }
        }

        // Check if activity party (from/to fields)
        if (attr.LogicalName.toLowerCase().includes('party') || 
            attr.LogicalName.toLowerCase().includes('from') || 
            attr.LogicalName.toLowerCase().includes('to')) {
          attribute.IsActivityParty = true;
        }

        return attribute;
      })
    );

    // Sort attributes
    attributes.sort((a, b) => (a.DisplayName || a.LogicalName).localeCompare(b.DisplayName || b.LogicalName));

    // Fetch all relationship types
    const [oneToMany, manyToOne, manyToMany] = await Promise.all([
      this.getRelationships(entityLogicalName, 'OneToMany').catch(() => []),
      this.getRelationships(entityLogicalName, 'ManyToOne').catch(() => []),
      this.getManyToManyRelationships(entityLogicalName).catch(() => [])
    ]);

    return {
      LogicalName: entityData.LogicalName,
      DisplayName: entityData.DisplayName?.UserLocalizedLabel?.Label || entityData.LogicalName,
      EntitySetName: entityData.EntitySetName,
      PrimaryIdAttribute: entityData.PrimaryIdAttribute,
      PrimaryNameAttribute: entityData.PrimaryNameAttribute,
      Description: entityData.Description?.UserLocalizedLabel?.Label,
      Attributes: attributes,
      OneToManyRelationships: oneToMany,
      ManyToOneRelationships: manyToOne,
      ManyToManyRelationships: manyToMany
    };
  }

  async getManyToManyRelationships(logicalName: string): Promise<RelationshipMetadata[]> {
    const url = `${this.baseUrl}/EntityDefinitions(LogicalName='${logicalName}')/ManyToManyRelationships`;
    const data = await this.fetchJson(url);

    return data.value.map((r: any) => ({
      SchemaName: r.SchemaName,
      ReferencingEntity: r.Entity1LogicalName,
      ReferencedEntity: r.Entity2LogicalName,
      ReferencingAttribute: '', // ManyToMany doesn't have a single referencing attribute
      ReferencingEntityNavigationPropertyName: r.Entity1NavigationPropertyName || '',
      ReferencedEntityNavigationPropertyName: r.Entity2NavigationPropertyName || '',
      RelationshipType: 'ManyToMany' as const
    }));
  }
}
