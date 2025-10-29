export class D365Helper {
  private overlayElements: HTMLElement[] = [];
  private requestCounter = 0;

  constructor() {
    // Communication happens via custom events with injected script
  }

  // Send request to injected script and wait for response
  private async sendRequest(action: string, data?: any): Promise<any> {
    return new Promise((resolve, reject) => {
      const requestId = `req_${this.requestCounter++}_${Date.now()}`;

      const timeout = setTimeout(() => {
        window.removeEventListener('D365_HELPER_RESPONSE', responseHandler);
        reject(new Error('Request timeout'));
      }, 5000);

      const responseHandler = (event: any) => {
        const response = event.detail;
        if (response.requestId === requestId) {
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

      window.dispatchEvent(new CustomEvent('D365_HELPER_REQUEST', {
        detail: { action, data, requestId }
      }));
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
      const controlInfo = await this.sendRequest('GET_CONTROL_INFO');

      console.log('D365 Helper: Creating overlays for', controlInfo.length, 'controls');

      controlInfo.forEach((info: any) => {
        try {
          // Try multiple ways to find the element
          let controlElement = document.getElementById(info.controlName);

          if (!controlElement) {
            // Try with data-id attribute
            controlElement = document.querySelector(`[data-id="${info.controlName}"]`);
          }

          if (!controlElement) {
            // Try partial ID match
            controlElement = document.querySelector(`[id*="${info.controlName}"]`);
          }

          if (controlElement) {
            // Try multiple ways to find a suitable container
            let container = controlElement.closest('[data-id]') ||
                          controlElement.closest('.control-container') ||
                          controlElement.closest('[data-control-name]') ||
                          controlElement.closest('div[role="group"]') ||
                          controlElement.parentElement;

            if (container) {
              const overlay = document.createElement('div');
              overlay.className = 'd365-schema-overlay';
              overlay.textContent = info.schemaName;
              overlay.title = `Schema Name: ${info.schemaName}\nLabel: ${info.label}\nClick to copy`;
              overlay.style.cssText = `
                position: absolute;
                top: 0;
                left: 0;
                background: rgba(0, 120, 212, 0.9);
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

              const parentElement = container as HTMLElement;
              const originalPosition = window.getComputedStyle(parentElement).position;
              if (originalPosition === 'static') {
                parentElement.style.position = 'relative';
              }

              parentElement.appendChild(overlay);
              this.overlayElements.push(overlay);
              console.debug('D365 Helper: Created overlay for', info.schemaName);
            } else {
              console.debug('D365 Helper: No container found for', info.controlName);
            }
          } else {
            console.debug('D365 Helper: Element not found for', info.controlName);
          }
        } catch (error) {
          console.debug('D365 Helper: Error creating overlay for', info.controlName, error);
        }
      });

      console.log('D365 Helper: Created', this.overlayElements.length, 'overlays');
    } catch (error) {
      console.error('Error showing schema overlay:', error);
    }
  }

  // Hide schema name overlay
  private hideSchemaOverlay(): void {
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
}
