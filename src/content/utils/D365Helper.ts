export class D365Helper {
  private overlayElements: HTMLElement[] = [];
  private requestCounter = 0;
  private headerObserver: MutationObserver | null = null;
  private headerFieldsInfo: any[] = [];
  private overlayColor: string = '#4bbf0d';

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
            if (!silent) {
              console.debug('[D365 Helper] Ignoring response from old injected script (no version).');
            }
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

      console.log('D365 Helper: Creating overlays for', controlInfo.length, 'controls');

      const processedContainers = new Set<HTMLElement>();
      const processedSchemaNames = new Set<string>();
      
      // Store unfound fields for later (when they become visible, e.g., header flyout opens)
      const unfoundFields: any[] = [];
      this.overlayColor = overlayColor;

      controlInfo.forEach((info: any) => {
        try {
          // Skip if we already created an overlay for this schema name
          if (processedSchemaNames.has(info.schemaName)) {
            console.debug('D365 Helper: Schema name already processed:', info.schemaName);
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
                console.debug('D365 Helper: Container already processed for', info.schemaName);
                return;
              }

              // Double-check: Don't add overlays to lookup value containers
              const parentClasses = parentElement.className;
              if (parentClasses.includes('Inline-Value') ||
                  parentClasses.includes('Inline-Item') ||
                  parentClasses.includes('lookupValue') ||
                  parentElement.querySelector('.ms-crm-Inline-Value, .ms-crm-Inline-Item')) {
                console.debug('D365 Helper: Skipping lookup value container for', info.schemaName);
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
              console.debug('D365 Helper: Created overlay for', info.schemaName);
            } else {
              console.debug('D365 Helper: No container found for', info.controlName);
              unfoundFields.push(info);
            }
          } else {
            console.debug('D365 Helper: Element not found for', info.controlName);
            unfoundFields.push(info);
          }
        } catch (error) {
          console.debug('D365 Helper: Error creating overlay for', info.controlName, error);
        }
      });

      console.log('D365 Helper: Created', this.overlayElements.length, 'overlays');
      console.log('D365 Helper:', unfoundFields.length, 'fields not found (may be in header flyout)');
      
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

      console.log('D365 Helper: Checking for', this.headerFieldsInfo.length, 'unfound fields');
      
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
          console.log('D365 Helper: Created overlay for previously unfound field:', info.schemaName);
        } catch (error) {
          stillUnfound.push(info);
          console.debug('D365 Helper: Error creating overlay for unfound field:', info.schemaName, error);
        }
      });

      // Update the list of still-unfound fields
      this.headerFieldsInfo = stillUnfound;
      
      if (stillUnfound.length === 0 && this.headerObserver) {
        console.log('D365 Helper: All fields found, disconnecting observer');
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

    console.log('D365 Helper: Observer set up for', this.headerFieldsInfo.length, 'unfound fields');
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
      return await this.sendRequest('GET_AUDIT_HISTORY');
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
}

