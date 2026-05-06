export interface EntityInfo {
  LogicalName: string;
  DisplayName: string;
  EntitySetName: string;
  PrimaryIdAttribute: string;
  PrimaryNameAttribute: string | null;
}

export class D365Helper {
  private overlayElements: HTMLElement[] = [];
  private requestCounter = 0;
  private headerObserver: MutationObserver | null = null;
  private headerFieldsInfo: any[] = [];
  private overlayColor: string = '#4bbf0d';
  private entityCache: EntityInfo[] | null = null;

  constructor() {
    // Communication happens via custom events with injected script
  }

  // Send request to injected script and wait for response
  private async sendRequest(
    action: string,
    data?: any,
    options?: { timeoutMs?: number; silent?: boolean }
  ): Promise<any> {
    return new Promise((resolve, reject) => {
      const requestId = `req_${this.requestCounter++}_${Date.now()}`;
      const timeoutMs = options?.timeoutMs ?? 5000;
      const silent = options?.silent ?? false;

      const timeout = setTimeout(() => {
        window.removeEventListener('D365_HELPER_RESPONSE', responseHandler);
        // Avoid noisy console errors for expected timeouts (e.g., during extension reloads)
        if (!silent) {
          console.warn('[D365 Helper] Request timeout:', action);
        }
        reject(new Error('Request timeout'));
      }, timeoutMs);

      const responseHandler = (event: any) => {
        const response = event.detail;
        if (response.requestId === requestId) {
          // Ignore responses from old scripts that don't include a version marker.
          // This prevents cached scripts from racing and causing false errors.
          if (!response._scriptVersion) {
            return; // Keep waiting for a versioned response
          }

          clearTimeout(timeout);
          window.removeEventListener('D365_HELPER_RESPONSE', responseHandler);

          if (response.success) {
            resolve(response.result);
          } else {
            reject(new Error(response.error));
          }
        }
      };

      window.addEventListener('D365_HELPER_RESPONSE', responseHandler);

      window.dispatchEvent(
        new CustomEvent('D365_HELPER_REQUEST', {
          detail: { action, data, requestId }
        })
      );
    });
  }

  // Get current record ID
  async getRecordId(): Promise<string | null> {
    try {
      return await this.sendRequest('GET_RECORD_ID');
    } catch {
      return null;
    }
  }

  // Get entity name
  async getEntityName(): Promise<string | null> {
    try {
      return await this.sendRequest('GET_ENTITY_NAME');
    } catch {
      return null;
    }
  }

  // Get organization URL
  getOrgUrl(): string {
    return window.location.origin;
  }

  // Get Web API URL for current record
  async getWebAPIUrl(): Promise<string | null> {
    const entityName = await this.getEntityName();
    const recordId = await this.getRecordId();

    if (!entityName || !recordId) return null;

    const orgUrl = this.getOrgUrl();
    const apiUrl = `${orgUrl}/api/data/v9.2/${this.getEntitySetName(entityName)}(${recordId})`;

    // Open in our custom viewer
    return chrome.runtime.getURL(`webapi-viewer.html?url=${encodeURIComponent(apiUrl)}`);
  }

  // Get Query Builder URL
  getQueryBuilderUrl(): string {
    const orgUrl = this.getOrgUrl();
    return chrome.runtime.getURL(`query-builder.html?orgUrl=${encodeURIComponent(orgUrl)}`);
  }

  // Get entity set name (pluralized logical name)
  private getEntitySetName(entityName: string): string {
    // Common irregular plurals
    const irregulars: { [key: string]: string } = {
      'opportunity': 'opportunities',
      'territory': 'territories',
      'currency': 'currencies',
      'activity': 'activities',
      'task': 'tasks'
    };

    if (irregulars[entityName]) {
      return irregulars[entityName];
    }

    // Simple pluralization
    if (entityName.endsWith('y')) {
      return entityName.slice(0, -1) + 'ies';
    } else if (entityName.endsWith('s')) {
      return entityName + 'es';
    } else {
      return entityName + 's';
    }
  }

  // Rollback a field to a previous value via WebAPI PATCH
  async rollbackFields(
    changes: { fieldName: string; oldValue: string }[],
    skipPlugins: boolean = false
  ): Promise<boolean> {
    const entityName = await this.getEntityName();
    const recordId = await this.getRecordId();

    if (!entityName || !recordId) {
      throw new Error('Could not determine current record context');
    }

    const orgUrl = this.getOrgUrl();
    const allEntities = await this.getAllEntities().catch(() => [] as EntityInfo[]);
    const entityByLogical = new Map<string, EntityInfo>();
    allEntities.forEach((entity) => entityByLogical.set(entity.LogicalName.toLowerCase(), entity));

    const sourceEntity = entityByLogical.get(entityName.toLowerCase());
    const entitySetName = sourceEntity?.EntitySetName || this.getEntitySetName(entityName);
    const apiUrl = `${orgUrl}/api/data/v9.2/${entitySetName}(${recordId})`;

    const escapeODataLiteral = (value: string): string => value.replace(/'/g, "''");

    const toNormalized = (value?: string | null): string =>
      String(value ?? '').trim().toLowerCase();

    const isEmptyAuditValue = (value?: string | null): boolean => {
      const normalized = toNormalized(value);
      return (
        normalized === '' ||
        normalized === '(empty)' ||
        normalized === 'empty' ||
        normalized === 'null'
      );
    };

    const extractGuid = (raw?: string | null): string | null => {
      if (!raw) return null;
      const match = raw.match(/[0-9a-fA-F]{8}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{12}/);
      if (!match) return null;
      const compact = match[0].replace(/[{}-]/g, '').toLowerCase();
      if (compact.length !== 32) return null;
      return `${compact.slice(0, 8)}-${compact.slice(8, 12)}-${compact.slice(12, 16)}-${compact.slice(16, 20)}-${compact.slice(20)}`;
    };

    const parseNumberValue = (raw: string): number | null => {
      const normalized = raw.replace(/,/g, '').replace(/[^\d.+-]/g, '').trim();
      if (!normalized) return null;
      const value = Number(normalized);
      return Number.isFinite(value) ? value : null;
    };

    interface OptionValue {
      value: number;
      label: string;
    }

    interface OptionAttribute {
      attributeType: string;
      isMultiSelect: boolean;
      options: OptionValue[];
    }

    interface LookupRelationship {
      navigationPropertyName: string;
      referencedEntity: string;
    }

    interface LookupTargetEntity {
      logicalName: string;
      entitySetName: string;
      primaryIdAttribute: string;
      primaryNameAttribute: string | null;
    }

    const optionAttributes = new Map<string, OptionAttribute>();

    try {
      const optionData = await this.getOptionSets();
      const attributes = Array.isArray(optionData?.attributes) ? optionData.attributes : [];
      attributes.forEach((attribute: any) => {
        const logicalName =
          typeof attribute?.logicalName === 'string' ? attribute.logicalName : '';
        if (!logicalName) return;

        const optionsRaw = Array.isArray(attribute?.options) ? attribute.options : [];
        const options: OptionValue[] = optionsRaw
          .map((option: any) => ({
            value: Number(option?.value),
            label: typeof option?.label === 'string' ? option.label : '',
          }))
          .filter((option: OptionValue) => Number.isFinite(option.value));

        optionAttributes.set(logicalName.toLowerCase(), {
          attributeType: String(attribute?.attributeType || ''),
          isMultiSelect: Boolean(attribute?.isMultiSelect),
          options,
        });
      });
    } catch (error) {
      // Option metadata helps convert labels, but rollback can still proceed without it.
    }

    const mapOptionLabelToValue = (fieldName: string, rawValue: string): number | null => {
      const optionInfo = optionAttributes.get(fieldName.toLowerCase());
      if (!optionInfo) return null;

      const numeric = Number(rawValue);
      if (Number.isFinite(numeric)) return numeric;

      const normalized = toNormalized(rawValue);
      const match = optionInfo.options.find(
        (option) => toNormalized(option.label) === normalized
      );
      return match ? match.value : null;
    };

    const attributeTypeCache = new Map<string, string>();
    const relationshipCache = new Map<string, LookupRelationship[]>();
    const targetEntityCache = new Map<string, LookupTargetEntity | null>();

    const getAttributeType = async (fieldName: string): Promise<string> => {
      const cacheKey = fieldName.toLowerCase();
      const cached = attributeTypeCache.get(cacheKey);
      if (cached) return cached;

      const metadataUrl =
        `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${escapeODataLiteral(entityName)}')` +
        `/Attributes(LogicalName='${escapeODataLiteral(fieldName)}')?$select=LogicalName,AttributeType`;

      const response = await fetch(metadataUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        credentials: 'include',
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(
          `Unable to read metadata for ${fieldName} (${response.status}): ${errorText}`
        );
      }

      const metadata = await response.json();
      const attributeType =
        typeof metadata?.AttributeType === 'string' ? metadata.AttributeType : 'String';

      attributeTypeCache.set(cacheKey, attributeType);
      return attributeType;
    };

    const getLookupRelationships = async (fieldName: string): Promise<LookupRelationship[]> => {
      const cacheKey = fieldName.toLowerCase();
      const cached = relationshipCache.get(cacheKey);
      if (cached) return cached;

      const relationshipsUrl =
        `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${escapeODataLiteral(entityName)}')` +
        `/ManyToOneRelationships?$select=ReferencingEntityNavigationPropertyName,ReferencedEntity` +
        `&$filter=ReferencingAttribute eq '${escapeODataLiteral(fieldName)}'`;

      const response = await fetch(relationshipsUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        credentials: 'include',
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(
          `Unable to resolve lookup relationship for ${fieldName} (${response.status}): ${errorText}`
        );
      }

      const json = await response.json();
      const relationships: LookupRelationship[] = Array.isArray(json?.value)
        ? json.value
            .map((item: any) => ({
              navigationPropertyName:
                typeof item?.ReferencingEntityNavigationPropertyName === 'string'
                  ? item.ReferencingEntityNavigationPropertyName
                  : '',
              referencedEntity:
                typeof item?.ReferencedEntity === 'string' ? item.ReferencedEntity : '',
            }))
            .filter(
              (item: LookupRelationship) =>
                item.navigationPropertyName.length > 0 && item.referencedEntity.length > 0
            )
        : [];

      relationshipCache.set(cacheKey, relationships);
      return relationships;
    };

    const getTargetEntityInfo = async (
      logicalName: string
    ): Promise<LookupTargetEntity | null> => {
      const cacheKey = logicalName.toLowerCase();
      if (targetEntityCache.has(cacheKey)) {
        return targetEntityCache.get(cacheKey) || null;
      }

      const fromCache = entityByLogical.get(cacheKey);
      if (fromCache) {
        const info: LookupTargetEntity = {
          logicalName: fromCache.LogicalName,
          entitySetName: fromCache.EntitySetName,
          primaryIdAttribute: fromCache.PrimaryIdAttribute,
          primaryNameAttribute: fromCache.PrimaryNameAttribute,
        };
        targetEntityCache.set(cacheKey, info);
        return info;
      }

      const entityUrl =
        `${orgUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${escapeODataLiteral(logicalName)}')` +
        `?$select=LogicalName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute`;

      const response = await fetch(entityUrl, {
        method: 'GET',
        headers: {
          Accept: 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        credentials: 'include',
      });

      if (!response.ok) {
        targetEntityCache.set(cacheKey, null);
        return null;
      }

      const json = await response.json();
      const entitySet =
        typeof json?.EntitySetName === 'string' ? json.EntitySetName : '';
      const idAttribute =
        typeof json?.PrimaryIdAttribute === 'string' ? json.PrimaryIdAttribute : '';
      if (!entitySet || !idAttribute) {
        targetEntityCache.set(cacheKey, null);
        return null;
      }

      const info: LookupTargetEntity = {
        logicalName: typeof json?.LogicalName === 'string' ? json.LogicalName : logicalName,
        entitySetName: entitySet,
        primaryIdAttribute: idAttribute,
        primaryNameAttribute:
          typeof json?.PrimaryNameAttribute === 'string' ? json.PrimaryNameAttribute : null,
      };
      targetEntityCache.set(cacheKey, info);
      return info;
    };

    const resolveLookupValue = async (
      fieldName: string,
      rawValue: string
    ): Promise<{ navigationPropertyName: string; entitySetName: string; recordId: string }> => {
      const relationships = await getLookupRelationships(fieldName);
      if (relationships.length === 0) {
        throw new Error(`No lookup relationship metadata found for ${fieldName}.`);
      }

      const guid = extractGuid(rawValue);

      if (guid) {
        if (relationships.length === 1) {
          const relationship = relationships[0];
          const target = await getTargetEntityInfo(relationship.referencedEntity);
          if (!target) {
            throw new Error(`Could not resolve target entity metadata for ${fieldName}.`);
          }
          return {
            navigationPropertyName: relationship.navigationPropertyName,
            entitySetName: target.entitySetName,
            recordId: guid,
          };
        }

        // For polymorphic lookups, verify which target table contains this GUID.
        for (const relationship of relationships) {
          const target = await getTargetEntityInfo(relationship.referencedEntity);
          if (!target) continue;

          const validateUrl =
            `${orgUrl}/api/data/v9.2/${target.entitySetName}(${guid})` +
            `?$select=${target.primaryIdAttribute}`;

          const validateResponse = await fetch(validateUrl, {
            method: 'GET',
            headers: {
              Accept: 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
            },
            credentials: 'include',
          });

          if (validateResponse.ok) {
            return {
              navigationPropertyName: relationship.navigationPropertyName,
              entitySetName: target.entitySetName,
              recordId: guid,
            };
          }
        }
      }

      // Resolve lookup by primary name when audit value is not a GUID.
      const lookupName = rawValue.trim();
      if (!lookupName) {
        throw new Error(`Lookup value for ${fieldName} is empty.`);
      }

      for (const relationship of relationships) {
        const target = await getTargetEntityInfo(relationship.referencedEntity);
        if (!target || !target.primaryNameAttribute) continue;

        const queryUrl =
          `${orgUrl}/api/data/v9.2/${target.entitySetName}?` +
          `$select=${target.primaryIdAttribute}&` +
          `$filter=${target.primaryNameAttribute} eq '${escapeODataLiteral(lookupName)}'&$top=2`;

        const response = await fetch(queryUrl, {
          method: 'GET',
          headers: {
            Accept: 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
          },
          credentials: 'include',
        });

        if (!response.ok) continue;

        const json = await response.json();
        const values = Array.isArray(json?.value) ? json.value : [];
        if (values.length === 0) continue;

        const id = values[0]?.[target.primaryIdAttribute];
        if (typeof id === 'string' && id.trim()) {
          return {
            navigationPropertyName: relationship.navigationPropertyName,
            entitySetName: target.entitySetName,
            recordId: id.replace(/[{}]/g, ''),
          };
        }
      }

      throw new Error(
        `Could not resolve lookup value "${rawValue}" for ${fieldName}. ` +
          `Use a GUID value for reliable rollback.`
      );
    };

    // Build the PATCH payload
    const payload: Record<string, any> = {};
    for (const change of changes) {
      const fieldName = String(change.fieldName || '').trim();
      if (!fieldName) continue;

      const rawValue = String(change.oldValue ?? '');
      const attributeType = (await getAttributeType(fieldName)).toLowerCase();

      if (attributeType === 'lookup' || attributeType === 'customer' || attributeType === 'owner') {
        const relationships = await getLookupRelationships(fieldName);
        if (relationships.length === 0) {
          throw new Error(`No lookup relationship metadata found for ${fieldName}.`);
        }

        if (isEmptyAuditValue(rawValue)) {
          // For polymorphic lookups, clear all related navigation properties.
          relationships.forEach((relationship) => {
            payload[relationship.navigationPropertyName] = null;
          });
          continue;
        }

        const resolved = await resolveLookupValue(fieldName, rawValue);
        payload[`${resolved.navigationPropertyName}@odata.bind`] =
          `/${resolved.entitySetName}(${resolved.recordId})`;
        continue;
      }

      if (isEmptyAuditValue(rawValue)) {
        payload[fieldName] = null;
        continue;
      }

      if (
        attributeType === 'picklist' ||
        attributeType === 'state' ||
        attributeType === 'status'
      ) {
        const mapped = mapOptionLabelToValue(fieldName, rawValue);
        if (mapped !== null) {
          payload[fieldName] = Math.trunc(mapped);
          continue;
        }

        const numeric = parseNumberValue(rawValue);
        if (numeric === null) {
          throw new Error(`Could not parse option value "${rawValue}" for ${fieldName}.`);
        }
        payload[fieldName] = Math.trunc(numeric);
        continue;
      }

      if (attributeType === 'multiselectpicklist') {
        const parts = rawValue
          .split(/[;,]/)
          .map((part) => part.trim())
          .filter(Boolean);

        const values: number[] = [];
        parts.forEach((part) => {
          const mapped = mapOptionLabelToValue(fieldName, part);
          if (mapped !== null) {
            values.push(Math.trunc(mapped));
            return;
          }

          const numeric = parseNumberValue(part);
          if (numeric !== null) {
            values.push(Math.trunc(numeric));
          }
        });

        if (values.length === 0) {
          throw new Error(
            `Could not parse multi-select option values "${rawValue}" for ${fieldName}.`
          );
        }

        payload[fieldName] = Array.from(new Set(values)).join(',');
        continue;
      }

      if (attributeType === 'boolean') {
        const normalized = toNormalized(rawValue);
        if (['true', '1', 'yes', 'y'].includes(normalized)) {
          payload[fieldName] = true;
          continue;
        }
        if (['false', '0', 'no', 'n'].includes(normalized)) {
          payload[fieldName] = false;
          continue;
        }

        const mapped = mapOptionLabelToValue(fieldName, rawValue);
        if (mapped !== null) {
          payload[fieldName] = mapped !== 0;
          continue;
        }

        throw new Error(`Could not parse boolean value "${rawValue}" for ${fieldName}.`);
      }

      if (attributeType === 'datetime') {
        const date = new Date(rawValue);
        if (!Number.isFinite(date.getTime())) {
          throw new Error(`Could not parse date value "${rawValue}" for ${fieldName}.`);
        }
        payload[fieldName] = date.toISOString();
        continue;
      }

      if (attributeType === 'integer' || attributeType === 'bigint') {
        const numeric = parseNumberValue(rawValue);
        if (numeric === null) {
          throw new Error(`Could not parse integer value "${rawValue}" for ${fieldName}.`);
        }
        payload[fieldName] = Math.trunc(numeric);
        continue;
      }

      if (
        attributeType === 'decimal' ||
        attributeType === 'double' ||
        attributeType === 'money'
      ) {
        const numeric = parseNumberValue(rawValue);
        if (numeric === null) {
          throw new Error(`Could not parse numeric value "${rawValue}" for ${fieldName}.`);
        }
        payload[fieldName] = numeric;
        continue;
      }

      if (attributeType === 'uniqueidentifier') {
        payload[fieldName] = extractGuid(rawValue) || rawValue.trim();
        continue;
      }

      // Default fallback: send as string value.
      payload[fieldName] = rawValue;
    }

    const headers: Record<string, string> = {
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      'OData-MaxVersion': '4.0',
      'OData-Version': '4.0',
    };

    if (skipPlugins) {
      headers['MSCRM.SuppressCallbackRegistrationExpanderJob'] = 'true';
      headers['MSCRM.BypassCustomPluginExecution'] = 'true';
    }

    const response = await fetch(apiUrl, {
      method: 'PATCH',
      headers,
      credentials: 'include',
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Rollback failed (${response.status}): ${errorText}`);
    }

    return true;
  }

  // Get all entities with caching
  async getAllEntities(): Promise<EntityInfo[]> {
    if (this.entityCache) return this.entityCache;

    const orgUrl = this.getOrgUrl();
    const url = `${orgUrl}/api/data/v9.2/EntityDefinitions?$select=LogicalName,DisplayName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute`;

    try {
      const response = await fetch(url, {
        headers: {
          'Accept': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        credentials: 'include',
      });

      if (!response.ok) {
        throw new Error(`Failed to load entities (${response.status})`);
      }

      const data = await response.json();
      const entities: EntityInfo[] = (data.value || []).map((e: any) => ({
        LogicalName: e.LogicalName,
        DisplayName: e.DisplayName?.UserLocalizedLabel?.Label || e.LogicalName,
        EntitySetName: e.EntitySetName,
        PrimaryIdAttribute: e.PrimaryIdAttribute,
        PrimaryNameAttribute: e.PrimaryNameAttribute ?? null,
      }));

      entities.sort((a, b) =>
        (a.DisplayName || a.LogicalName).localeCompare(b.DisplayName || b.LogicalName)
      );

      this.entityCache = entities;
      return entities;
    } catch (err) {
      // Fallback without $select
      const fallbackUrl = `${orgUrl}/api/data/v9.2/EntityDefinitions`;
      const response = await fetch(fallbackUrl, {
        headers: {
          'Accept': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        credentials: 'include',
      });

      if (!response.ok) {
        throw new Error(`Failed to load entities (${response.status})`);
      }

      const data = await response.json();
      const entities: EntityInfo[] = (data.value || []).map((e: any) => ({
        LogicalName: e.LogicalName,
        DisplayName: e.DisplayName?.UserLocalizedLabel?.Label || e.LogicalName,
        EntitySetName: e.EntitySetName,
        PrimaryIdAttribute: e.PrimaryIdAttribute,
        PrimaryNameAttribute: e.PrimaryNameAttribute ?? null,
      }));

      entities.sort((a, b) =>
        (a.DisplayName || a.LogicalName).localeCompare(b.DisplayName || b.LogicalName)
      );

      this.entityCache = entities;
      return entities;
    }
  }

  // Get form editor URL
  async getFormEditorUrl(): Promise<string | null> {
    try {
      const entityName = await this.getEntityName();
      const formId = await this.sendRequest('GET_FORM_ID');
      const orgUrl = this.getOrgUrl();

      return `${orgUrl}/main.aspx?appid=&pagetype=formeditor&formid=${formId}&entitytype=${entityName}`;
    } catch {
      return null;
    }
  }

  // Get environment ID by querying the Web API
  async getEnvironmentId(): Promise<string | null> {
    try {
      const orgUrl = this.getOrgUrl();
      const response = await fetch(`${orgUrl}/api/data/v9.2/RetrieveCurrentOrganization()`, {
        method: 'GET',
        headers: {
          'Accept': 'application/json',
          'OData-MaxVersion': '4.0',
          'OData-Version': '4.0',
        },
        credentials: 'include'
      });

      if (response.ok) {
        const data = await response.json();
        // EnvironmentId is preferred, but fall back to Id/OrganizationId if needed
        const rawEnvironmentId = data?.EnvironmentId ?? data?.Id ?? data?.OrganizationId;

        if (rawEnvironmentId) {
          return String(rawEnvironmentId).replace(/[{}]/g, '');
        }

        console.warn('Environment identifier not found in RetrieveCurrentOrganization response', data);
      }
      return null;
    } catch (error) {
      console.error('Error getting environment ID:', error);
      return null;
    }
  }

  // Get solutions page URL (Power Apps maker portal)
  async getSolutionsUrl(): Promise<string | null> {
    try {
      const environmentId = await this.getEnvironmentId();
      if (environmentId) {
        // Add timestamp to prevent caching
        const timestamp = Date.now();
        return `https://make.powerapps.com/environments/${environmentId}/solutions?_=${timestamp}`;
      }
      return `https://make.powerapps.com/`;
    } catch {
      return `https://make.powerapps.com/`;
    }
  }

  // Get Power Platform admin center URL
  getAdminCenterUrl(): string {
    return `https://admin.powerplatform.microsoft.com/manage/environments`;
  }

  // Retrieve plugin trace logs
  async getPluginTraceLogs(limit: number = 20): Promise<any> {
    try {
      return await this.sendRequest('GET_PLUGIN_TRACE_LOGS', { top: limit });
    } catch (error) {
      console.error('Error retrieving plugin trace logs:', error);
      throw error;
    }
  }

  // Toggle all fields visibility
  async toggleAllFields(show: boolean): Promise<void> {
    try {
      await this.sendRequest('TOGGLE_FIELDS', { show });
    } catch (error) {
      console.error('Error toggling fields:', error);
      throw error;
    }
  }

  // Toggle all sections visibility
  async toggleAllSections(show: boolean): Promise<void> {
    try {
      await this.sendRequest('TOGGLE_SECTIONS', { show });
    } catch (error) {
      console.error('Error toggling sections:', error);
      throw error;
    }
  }

  // Toggle blur on all field values
  async toggleBlurFields(blur: boolean): Promise<void> {
    try {
      await this.sendRequest('TOGGLE_BLUR_FIELDS', { blur });
    } catch (error) {
      console.error('Error toggling field blur:', error);
      throw error;
    }
  }

  // Get all schema names
  async getAllSchemaNames(): Promise<string[]> {
    try {
      return await this.sendRequest('GET_SCHEMA_NAMES');
    } catch (error) {
      console.error('Error getting schema names:', error);
      return [];
    }
  }

  // Unlock readonly fields
  async unlockFields(): Promise<number> {
    try {
      const result = await this.sendRequest('UNLOCK_FIELDS');
      return result.unlockedCount;
    } catch (error) {
      console.error('Error unlocking fields:', error);
      throw error;
    }
  }

  // Auto-fill form with sample data
  async autoFillForm(): Promise<number> {
    try {
      const result = await this.sendRequest('AUTO_FILL_FORM');
      return result.filledCount;
    } catch (error) {
      console.error('Error auto-filling form:', error);
      throw error;
    }
  }

  // Disable field requirements
  async disableFieldRequirements(): Promise<number> {
    try {
      const result = await this.sendRequest('DISABLE_REQUIRED_FIELDS');
      return result.disabledCount;
    } catch (error) {
      console.error('Error disabling field requirements:', error);
      throw error;
    }
  }

  // Retrieve option sets for current form
  async getOptionSets(): Promise<any> {
    try {
      return await this.sendRequest('GET_OPTION_SETS');
    } catch (error) {
      console.error('Error retrieving option sets:', error);
      throw error;
    }
  }

  // Toggle schema name overlay
  async toggleSchemaOverlay(show: boolean): Promise<void> {
    if (show) {
      await this.showSchemaOverlay();
    } else {
      this.hideSchemaOverlay();
    }
  }

  // Show schema name overlay
  private async showSchemaOverlay(): Promise<void> {
    this.hideSchemaOverlay(); // Clear any existing overlays

    try {
      // Get schema overlay color from settings
      const settings = await chrome.storage.sync.get(['schemaOverlayColor']);
      // Force set to default if not already set
      if (!settings.schemaOverlayColor || settings.schemaOverlayColor !== '#4bbf0d') {
        chrome.storage.sync.set({ schemaOverlayColor: '#4bbf0d' });
      }
      const overlayColor = settings.schemaOverlayColor || '#4bbf0d';

      // Convert hex to rgba
      const hexToRgba = (hex: string, alpha: number = 0.9): string => {
        const r = parseInt(hex.slice(1, 3), 16);
        const g = parseInt(hex.slice(3, 5), 16);
        const b = parseInt(hex.slice(5, 7), 16);
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
      };

      const bgColor = hexToRgba(overlayColor, 0.9);

      const controlInfo = await this.sendRequest('GET_CONTROL_INFO');

      const processedContainers = new Set<HTMLElement>();
      const processedSchemaNames = new Set<string>();
      
      // Store unfound fields for later (when they become visible, e.g., header flyout opens)
      const unfoundFields: any[] = [];
      this.overlayColor = overlayColor;

      controlInfo.forEach((info: any) => {
        try {
          // Skip if we already created an overlay for this schema name
          if (processedSchemaNames.has(info.schemaName)) {
            return;
          }

          // Try multiple ways to find the element
          let controlElement = document.getElementById(info.controlName);

          if (!controlElement) {
            // Try with data-id attribute
            controlElement = document.querySelector(`[data-id="${info.controlName}"]`) as HTMLElement;
          }

          if (!controlElement) {
            // Try partial ID match
            controlElement = document.querySelector(`[id*="${info.controlName}"]`) as HTMLElement;
          }

          // Additional search strategies for header fields and other cases
          if (!controlElement) {
            // Try finding by aria-label or aria-describedby that might contain the control name
            const ariaElements = document.querySelectorAll(`[aria-label*="${info.controlName}"], [aria-describedby*="${info.controlName}"]`);
            if (ariaElements.length > 0) {
              controlElement = ariaElements[0] as HTMLElement;
            }
          }

          if (!controlElement) {
            // Try finding in header section specifically - use multiple selectors
            const headerSelectors = [
              '[data-id="header"]',
              '.ms-crm-Form-Header',
              '[class*="header"]',
              '[class*="Header"]',
              '[id*="header"]',
              '[id*="Header"]',
              '[class*="form-header"]',
              '[class*="FormHeader"]'
            ];
            
            for (const headerSelector of headerSelectors) {
              const headerSection = document.querySelector(headerSelector);
              if (headerSection) {
                const headerControl = headerSection.querySelector(
                  `[data-id="${info.controlName}"], [id*="${info.controlName}"], [data-lp-id="${info.controlName}"], [data-control-name="${info.controlName}"]`
                );
                if (headerControl) {
                  controlElement = headerControl as HTMLElement;
                  break;
                }
              }
            }
          }

          // If still not found, try to find any element with the control name in various attributes
          if (!controlElement) {
            const fallbackSelectors = [
              `[data-lp-id="${info.controlName}"]`,
              `[name="${info.controlName}"]`,
              `input[id*="${info.controlName}"]`,
              `select[id*="${info.controlName}"]`,
              `textarea[id*="${info.controlName}"]`,
              `[data-control-name="${info.controlName}"]`
            ];
            
            for (const selector of fallbackSelectors) {
              const found = document.querySelector(selector);
              if (found) {
                controlElement = found as HTMLElement;
                break;
              }
            }
          }

          if (controlElement) {
            // Helper function to check if element is visible
            const isVisible = (elem: HTMLElement): boolean => {
              const style = window.getComputedStyle(elem);
              return style.display !== 'none' && 
                     style.visibility !== 'hidden' && 
                     style.opacity !== '0' &&
                     elem.offsetWidth > 0 && 
                     elem.offsetHeight > 0;
            };
            
            // Find the proper field container - prioritize smaller, visible containers
            let container: HTMLElement | null = null;

            // Strategy: Find the smallest visible container that contains the control
            const containerCandidates = [
              // Most specific - direct data-id container
              controlElement.closest(`[data-id="${info.controlName}"]`),
              controlElement.closest('[data-id]'),
              // Field/control specific containers
              controlElement.closest('[class*="field"][class*="container"]'),
              controlElement.closest('[class*="control"][class*="container"]'),
              controlElement.closest('[class*="Field"][class*="Container"]'),
              controlElement.closest('[class*="Control"][class*="Container"]'),
              // Generic field containers
              controlElement.closest('[class*="field"]'),
              controlElement.closest('[class*="Field"]'),
              controlElement.closest('[class*="control"]'),
              controlElement.closest('[class*="Control"]'),
              // Role-based containers
              controlElement.closest('div[role="group"]'),
              controlElement.closest('[data-control-name]'),
              // Header-specific containers (but only if they're not too large)
              controlElement.closest('.ms-crm-Form-Header'),
              controlElement.closest('[class*="header"]'),
              controlElement.closest('[class*="Header"]'),
              // Section containers
              controlElement.closest('.ms-crm-FormSection'),
              controlElement.closest('.ms-crm-FormBody'),
              // Parent elements
              controlElement.parentElement,
              controlElement.parentElement?.parentElement
            ].filter(c => c !== null) as HTMLElement[];

            // Find the smallest visible container that's not too large
            for (const candidate of containerCandidates) {
              if (!candidate || !isVisible(candidate)) continue;
              
              // Skip if container is too large (likely a page-level container)
              const rect = candidate.getBoundingClientRect();
              if (rect.width > window.innerWidth * 0.9 || rect.height > window.innerHeight * 0.9) {
                continue;
              }
              
              // Skip if it's a lookup value container
              const classList = candidate.classList.toString();
              const hasLookupClass = classList.includes('lookup') ||
                                    classList.includes('ms-crm-Inline-Value') ||
                                    classList.includes('ms-crm-Inline-Item');

              // Skip if it's inside a lookup value display
              const isInLookupValue = candidate.closest('.ms-crm-Inline-Value, .ms-crm-Inline-Item, [class*="lookupValue"]');

              if (!hasLookupClass && !isInLookupValue) {
                // Prefer smaller containers
                if (!container || 
                    (candidate.contains(controlElement) && 
                     candidate.getBoundingClientRect().width < container.getBoundingClientRect().width)) {
                  container = candidate;
                }
              }
            }

            // Fallback to parent if no suitable container found
            if (!container) {
              container = controlElement.parentElement;
            }

            if (container && isVisible(container)) {
              const parentElement = container;

              if (processedContainers.has(parentElement)) {
                return;
              }

              // Double-check: Don't add overlays to lookup value containers
              const parentClasses = parentElement.className;
              if (parentClasses.includes('Inline-Value') ||
                  parentClasses.includes('Inline-Item') ||
                  parentClasses.includes('lookupValue') ||
                  parentElement.querySelector('.ms-crm-Inline-Value, .ms-crm-Inline-Item')) {
                return;
              }

              processedContainers.add(parentElement);
              processedSchemaNames.add(info.schemaName);

              const overlay = document.createElement('div');
              overlay.className = 'd365-schema-overlay';
              overlay.textContent = info.schemaName;
              overlay.title = `Schema Name: ${info.schemaName}\nLabel: ${info.label}\nClick to copy`;
              overlay.style.cssText = `
                position: absolute;
                top: 0;
                left: 0;
                background: ${bgColor};
                color: white;
                padding: 2px 6px;
                font-size: 11px;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                border-radius: 0 0 4px 0;
                z-index: 99999;
                cursor: pointer;
                pointer-events: auto;
              `;

              overlay.addEventListener('click', async (e) => {
                e.stopPropagation();
                await navigator.clipboard.writeText(info.schemaName);
                overlay.textContent = '✓ Copied!';
                setTimeout(() => {
                  overlay.textContent = info.schemaName;
                }, 1000);
              });

              const originalPosition = window.getComputedStyle(parentElement).position;
              if (originalPosition === 'static') {
                parentElement.style.position = 'relative';
              }

              parentElement.appendChild(overlay);
              this.overlayElements.push(overlay);
            } else {
              unfoundFields.push(info);
            }
          } else {
            unfoundFields.push(info);
          }
        } catch (error) {
          // Overlay creation failed for this control; skip it.
        }
      });

      // Store unfound fields and set up observer
      if (unfoundFields.length > 0) {
        this.headerFieldsInfo = unfoundFields;
        this.setupHeaderFlyoutObserver(bgColor, processedContainers, processedSchemaNames);
      }
    } catch (error) {
      console.error('Error showing schema overlay:', error);
    }
  }

  // Set up observer to watch for unfound fields becoming visible (e.g., header flyout opening)
  private setupHeaderFlyoutObserver(bgColor: string, processedContainers: Set<HTMLElement>, processedSchemaNames: Set<string>): void {
    // Clean up existing observer
    if (this.headerObserver) {
      this.headerObserver.disconnect();
      this.headerObserver = null;
    }

    // Helper function to check if element is visible
    const isVisible = (elem: HTMLElement): boolean => {
      const style = window.getComputedStyle(elem);
      const rect = elem.getBoundingClientRect();
      return style.display !== 'none' && 
             style.visibility !== 'hidden' && 
             style.opacity !== '0' &&
             rect.width > 0 && 
             rect.height > 0;
    };

    // Function to try creating overlays for unfound fields
    const tryCreateOverlaysForUnfoundFields = () => {
      if (this.headerFieldsInfo.length === 0) return;

      const stillUnfound: any[] = [];

      this.headerFieldsInfo.forEach((info: any) => {
        try {
          if (processedSchemaNames.has(info.schemaName)) {
            return;
          }

          // Try to find the control element
          const selectors = [
            `[data-id="${info.controlName}"]`,
            `[id="${info.controlName}"]`,
            `[id*="${info.controlName}"]`,
            `[data-lp-id="${info.controlName}"]`,
            `[data-control-name="${info.controlName}"]`
          ];

          let controlElement: HTMLElement | null = null;
          for (const sel of selectors) {
            try {
              const found = document.querySelector(sel);
              if (found && isVisible(found as HTMLElement)) {
                controlElement = found as HTMLElement;
                break;
              }
            } catch (e) {
              // Selector might be invalid
            }
          }

          if (!controlElement) {
            stillUnfound.push(info);
            return;
          }

          // Find the smallest visible container
          const containerCandidates = [
            controlElement.closest(`[data-id="${info.controlName}"]`),
            controlElement.closest('[data-id]'),
            controlElement.closest('[class*="field"]'),
            controlElement.closest('[class*="Field"]'),
            controlElement.closest('[class*="control"]'),
            controlElement.closest('[class*="Control"]'),
            controlElement.closest('div[role="group"]'),
            controlElement.parentElement
          ].filter(c => c !== null) as HTMLElement[];

          let container: HTMLElement | null = null;
          for (const candidate of containerCandidates) {
            if (!candidate || !isVisible(candidate)) continue;
            
            // Skip if too large
            const rect = candidate.getBoundingClientRect();
            if (rect.width > window.innerWidth * 0.8 || rect.height > window.innerHeight * 0.8) {
              continue;
            }
            
            // Skip lookup containers
            const classes = candidate.className || '';
            if (classes.includes('Inline-Value') || classes.includes('Inline-Item') || classes.includes('lookupValue')) {
              continue;
            }

            container = candidate;
            break;
          }

          if (!container) {
            container = controlElement.parentElement;
          }

          if (!container || processedContainers.has(container)) {
            stillUnfound.push(info);
            return;
          }

          // Check visibility again
          if (!isVisible(container)) {
            stillUnfound.push(info);
            return;
          }

          processedContainers.add(container);
          processedSchemaNames.add(info.schemaName);

          const overlay = document.createElement('div');
          overlay.className = 'd365-schema-overlay d365-schema-overlay-header';
          overlay.textContent = info.schemaName;
          overlay.title = `Schema Name: ${info.schemaName}\nLabel: ${info.label}\nClick to copy`;
          overlay.style.cssText = `
            position: absolute;
            top: 0;
            left: 0;
            background: ${bgColor};
            color: white;
            padding: 2px 6px;
            font-size: 11px;
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            border-radius: 0 0 4px 0;
            z-index: 99999;
            cursor: pointer;
            pointer-events: auto;
          `;

          overlay.addEventListener('click', async (e) => {
            e.stopPropagation();
            await navigator.clipboard.writeText(info.schemaName);
            overlay.textContent = '✓ Copied!';
            setTimeout(() => {
              overlay.textContent = info.schemaName;
            }, 1000);
          });

          const originalPosition = window.getComputedStyle(container).position;
          if (originalPosition === 'static') {
            container.style.position = 'relative';
          }

          container.appendChild(overlay);
          this.overlayElements.push(overlay);
        } catch (error) {
          stillUnfound.push(info);
        }
      });

      // Update the list of still-unfound fields
      this.headerFieldsInfo = stillUnfound;
      
      if (stillUnfound.length === 0 && this.headerObserver) {
        this.headerObserver.disconnect();
        this.headerObserver = null;
      }
    };

    // Set up MutationObserver to watch for DOM changes
    this.headerObserver = new MutationObserver((mutations) => {
      // Debounce - only check after DOM settles
      setTimeout(tryCreateOverlaysForUnfoundFields, 150);
    });

    // Observe the document body for changes
    this.headerObserver.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ['class', 'style', 'aria-expanded', 'aria-hidden', 'hidden']
    });

    // Also set up periodic check (in case mutation observer misses something)
    const checkInterval = setInterval(() => {
      if (this.headerFieldsInfo.length === 0) {
        clearInterval(checkInterval);
        return;
      }
      tryCreateOverlaysForUnfoundFields();
    }, 1000);

    // Stop checking after 30 seconds
    setTimeout(() => {
      clearInterval(checkInterval);
    }, 30000);
  }

  // Hide schema name overlay
  private hideSchemaOverlay(): void {
    // Clean up header observer
    if (this.headerObserver) {
      this.headerObserver.disconnect();
      this.headerObserver = null;
    }
    
    // Clear header fields info
    this.headerFieldsInfo = [];
    
    // Remove all overlays from DOM
    this.overlayElements.forEach(overlay => {
      try {
        if (overlay.parentElement) {
          overlay.parentElement.removeChild(overlay);
        }
      } catch (error) {
        // Element might already be removed
      }
    });
    this.overlayElements = [];

    // Also remove any orphaned overlays that might exist in the DOM
    const orphanedOverlays = document.querySelectorAll('.d365-schema-overlay');
    orphanedOverlays.forEach(overlay => {
      try {
        overlay.remove();
      } catch (error) {
        // Ignore
      }
    });
  }

  // Get form libraries and event handlers
  async getFormLibraries(): Promise<any> {
    try {
      return await this.sendRequest('GET_FORM_LIBRARIES');
    } catch (error) {
      console.error('Error getting form libraries:', error);
      throw error;
    }
  }

  // Get OData fields metadata for current entity
  async getODataFields(): Promise<any> {
    try {
      return await this.sendRequest('GET_ODATA_FIELDS');
    } catch (error) {
      console.error('Error getting OData fields:', error);
      throw error;
    }
  }

  // Get audit history for current record
  async getAuditHistory(): Promise<any> {
    try {
      // Audit history can be slow in some environments; allow longer timeout.
      // This is a hard timeout, not a wait—if it returns sooner, we resolve immediately.
      return await this.sendRequest('GET_AUDIT_HISTORY', undefined, { timeoutMs: 120000 });
    } catch (error) {
      console.error('Error getting audit history:', error);
      throw error;
    }
  }

  // ===== IMPERSONATION METHODS =====

  // Get list of system users for impersonation selector
  async getSystemUsers(): Promise<any> {
    try {
      return await this.sendRequest('GET_SYSTEM_USERS');
    } catch (error) {
      console.error('Error getting system users:', error);
      throw error;
    }
  }

  // Set impersonation for a specific user
  async setImpersonation(userId: string, fullname: string, domainname: string): Promise<any> {
    try {
      return await this.sendRequest('SET_IMPERSONATION', { userId, fullname, domainname });
    } catch (error) {
      console.error('Error setting impersonation:', error);
      throw error;
    }
  }

  // Clear impersonation and return to original user
  async clearImpersonation(): Promise<any> {
    try {
      return await this.sendRequest('CLEAR_IMPERSONATION');
    } catch (error) {
      console.error('Error clearing impersonation:', error);
      throw error;
    }
  }

  // Check current impersonation status
  async getImpersonationStatus(): Promise<{ isImpersonating: boolean; user: any | null }> {
    try {
      return await this.sendRequest('GET_IMPERSONATION_STATUS', undefined, { silent: true, timeoutMs: 1500 });
    } catch (error) {
      return { isImpersonating: false, user: null };
    }
  }

  // ===== ACTIVE PROCESSES =====

  async getActiveProcesses(entityName?: string): Promise<any> {
    return await this.sendRequest('GET_ACTIVE_PROCESSES', { entityName }, { timeoutMs: 30000 });
  }

  async toggleProcess(id: string, activate: boolean): Promise<{ success: boolean; error?: string }> {
    return await this.sendRequest('TOGGLE_PROCESS', { id, activate }, { timeoutMs: 30000 });
  }

  // ===== PLUGIN STEPS =====

  async getPluginSteps(entityName?: string): Promise<any> {
    return await this.sendRequest('GET_PLUGIN_STEPS', { entityName }, { timeoutMs: 30000 });
  }

  async togglePluginStep(id: string, enable: boolean): Promise<{ success: boolean; error?: string }> {
    return await this.sendRequest('TOGGLE_PLUGIN_STEP', { id, enable }, { timeoutMs: 30000 });
  }

  async updatePluginStepImage(args: {
    id: string;
    name?: string;
    entityAlias?: string;
    attributes?: string;
    imageType?: number;
  }): Promise<{ success: boolean; error?: string }> {
    return await this.sendRequest('UPDATE_PLUGIN_STEP_IMAGE', args, { timeoutMs: 30000 });
  }

  async deletePluginStepImage(id: string): Promise<{ success: boolean; error?: string }> {
    return await this.sendRequest('DELETE_PLUGIN_STEP_IMAGE', { id }, { timeoutMs: 30000 });
  }

  async createPluginStepImage(args: {
    stepId: string;
    name: string;
    entityAlias: string;
    attributes: string;
    imageType: number;
    messagePropertyName?: string;
  }): Promise<{ success: boolean; id?: string; error?: string }> {
    return await this.sendRequest('CREATE_PLUGIN_STEP_IMAGE', args, { timeoutMs: 30000 });
  }

  // ===== PRIVILEGE DEBUGGER =====

  async getPrivilegeDebug(entityName: string, recordId: string): Promise<any> {
    return await this.sendRequest('GET_PRIVILEGE_DEBUG', { entityName, recordId }, { timeoutMs: 60000 });
  }
}

