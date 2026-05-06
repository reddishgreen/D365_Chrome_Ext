// This script runs in the page context and has access to window.Xrm
// It communicates with the content script via custom events

// ===== IMPERSONATION INTERCEPTION =====
// Intercept fetch and XMLHttpRequest to add MSCRMCallerID header when impersonation is active

// Helper to check if URL is a D365 Web API URL
function isD365ApiUrl(url: string | URL | Request): boolean {
  const urlString = typeof url === 'string' ? url : url instanceof URL ? url.href : url.url;
  return urlString.includes('/api/data/') || urlString.includes('/api/');
}

// Store original fetch
const originalFetch = window.fetch.bind(window);

// Override fetch to add impersonation header
(window as any).fetch = async function(input: RequestInfo | URL, init?: RequestInit): Promise<Response> {
  const impersonatedUser = (window as any).__d365ImpersonatedUser;
  
  if (impersonatedUser && isD365ApiUrl(input)) {
    // IMPORTANT: D365 often calls fetch(Request) with headers already set on the Request.
    // If we pass init.headers without including Request.headers, we can accidentally drop
    // Content-Type / OData headers, which breaks PATCH/POST (exactly the error you saw).
    const mergedHeaders = new Headers();

    // Preserve headers from Request input
    if (input instanceof Request) {
      input.headers.forEach((value, key) => {
        mergedHeaders.set(key, value);
      });
    }

    // Preserve/override headers from init (if provided)
    if (init?.headers) {
      const initHeaders = new Headers(init.headers as any);
      initHeaders.forEach((value, key) => {
        mergedHeaders.set(key, value);
      });
    }

    // Add MSCRMCallerID if not already present
    if (!mergedHeaders.has('MSCRMCallerID')) {
      mergedHeaders.set('MSCRMCallerID', impersonatedUser.systemuserid);
    }

    init = { ...(init || {}), headers: mergedHeaders };
  }
  
  return originalFetch(input, init);
};

// Store original XMLHttpRequest.open
const originalXHROpen = XMLHttpRequest.prototype.open;
const originalXHRSend = XMLHttpRequest.prototype.send;

// Track pending requests that need impersonation header
const xhrImpersonationMap = new WeakMap<XMLHttpRequest, boolean>();

// Override XMLHttpRequest.open to track API calls
XMLHttpRequest.prototype.open = function(method: string, url: string | URL, ...args: any[]) {
  const urlString = typeof url === 'string' ? url : url.href;
  if (isD365ApiUrl(urlString)) {
    xhrImpersonationMap.set(this, true);
  }
  return originalXHROpen.apply(this, [method, url, ...args] as any);
};

// Override XMLHttpRequest.send to add impersonation header
XMLHttpRequest.prototype.send = function(body?: Document | XMLHttpRequestBodyInit | null) {
  const impersonatedUser = (window as any).__d365ImpersonatedUser;
  
  if (impersonatedUser && xhrImpersonationMap.get(this)) {
    this.setRequestHeader('MSCRMCallerID', impersonatedUser.systemuserid);
  }
  
  return originalXHRSend.call(this, body);
};

// Version identifier for debugging
const INJECTED_SCRIPT_VERSION = '1.7.0-step-images';

// ===== END IMPERSONATION INTERCEPTION =====

// Store original visibility states
// Separate maps for TOGGLE_FIELDS and TOGGLE_SECTIONS to avoid conflicts
const originalFieldVisibility = new Map<string, boolean>();
const originalSectionVisibilityForFields = new Map<string, boolean>(); // Used by TOGGLE_FIELDS
const originalSectionVisibilityForSections = new Map<string, boolean>(); // Used by TOGGLE_SECTIONS

// Listen for requests from content script
window.addEventListener('D365_HELPER_REQUEST', async (event: any) => {
  const { action, data, requestId } = event.detail;
  
  try {
    let result: any = null;

    const Xrm = (window as any).Xrm;

    const requiresFormContext =
      action !== 'GET_PLUGIN_TRACE_LOGS' &&
      action !== 'GET_ENVIRONMENT_ID' &&
      action !== 'GET_SYSTEM_USERS' &&
      action !== 'SET_IMPERSONATION' &&
      action !== 'CLEAR_IMPERSONATION' &&
      action !== 'GET_IMPERSONATION_STATUS';

    if (requiresFormContext && (!Xrm || !Xrm.Page)) {
      throw new Error('Xrm.Page not available');
    }

    switch (action) {
      case 'GET_RECORD_ID':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }
        result = Xrm.Page.data.entity.getId().replace(/[{}]/g, '');
        break;

      case 'GET_ENTITY_NAME':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }
        result = Xrm.Page.data.entity.getEntityName();
        break;

      case 'GET_FORM_ID':
        result = Xrm.Page.ui.formSelector.getCurrentItem().getId();
        break;

      case 'GET_ENVIRONMENT_ID':
        // Get the organization ID which is the environment ID
        result = Xrm.Utility.getGlobalContext().organizationSettings.organizationId.replace(/[{}]/g, '');
        break;

      case 'TOGGLE_FIELDS':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }

        const attributes = Xrm.Page.data.entity.attributes.get();

        if (data.show) {
          // Clear any existing saved states and save current state before showing all
          originalFieldVisibility.clear();
          originalSectionVisibilityForFields.clear();

          // Build a control-to-section visibility map upfront
          // This maps each control name to its parent section's visibility state
          const controlSectionVisibility = new Map();
          const tabs = Xrm.Page.ui.tabs.get();

          // First pass: Save section visibility and build control-to-section map
          tabs.forEach((tab: any) => {
            const sections = tab.sections.get();
            sections.forEach((section: any) => {
              const sectionName = section.getName();
              const sectionVisible = section.getVisible();

              // Save original section visibility
              originalSectionVisibilityForFields.set(sectionName, sectionVisible);

              try {
                const sectionControls = section.controls.get();
                sectionControls.forEach((control: any) => {
                  const controlName = control.getName();
                  controlSectionVisibility.set(controlName, sectionVisible);
                });
              } catch (e) {
                // Silently ignore errors getting section controls
              }

              // Show all sections
              section.setVisible(true);
            });
          });

          // Second pass: Save field visibility states considering section visibility
          attributes.forEach((attribute: any) => {
            const controls = attribute.controls.get();
            controls.forEach((control: any) => {
              const controlName = control.getName();
              let actuallyVisible = control.getVisible();

              // Check if this control's parent section is hidden
              const sectionVisible = controlSectionVisibility.get(controlName);
              if (sectionVisible !== undefined) {
                // Field is only actually visible if both field AND section are visible
                actuallyVisible = actuallyVisible && sectionVisible;
              }

              originalFieldVisibility.set(controlName, actuallyVisible);
              control.setVisible(true);
            });
          });
        } else {
          // Restore original section visibility first
          const tabs = Xrm.Page.ui.tabs.get();
          tabs.forEach((tab: any) => {
            const sections = tab.sections.get();
            sections.forEach((section: any) => {
              const sectionName = section.getName();
              const originalState = originalSectionVisibilityForFields.get(sectionName);
              if (originalState !== undefined) {
                section.setVisible(originalState);
              }
            });
          });

          // Then restore field visibility
          attributes.forEach((attribute: any) => {
            const controls = attribute.controls.get();
            controls.forEach((control: any) => {
              const controlName = control.getName();
              const originalState = originalFieldVisibility.get(controlName);
              if (originalState !== undefined) {
                control.setVisible(originalState);
              }
            });
          });

          // Clear the saved states after restoration
          originalFieldVisibility.clear();
          originalSectionVisibilityForFields.clear();
        }

        result = { success: true };
        break;

      case 'TOGGLE_SECTIONS':
        if (!Xrm || !Xrm.Page || !Xrm.Page.ui) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }

        const tabs = Xrm.Page.ui.tabs.get();

        if (data.show) {
          // Clear any existing saved states and save current state before showing all
          originalSectionVisibilityForSections.clear();
          tabs.forEach((tab: any) => {
            const sections = tab.sections.get();
            sections.forEach((section: any) => {
              const sectionName = section.getName();
              originalSectionVisibilityForSections.set(sectionName, section.getVisible());
              section.setVisible(true);
            });
          });
        } else {
          // Restore original visibility state
          tabs.forEach((tab: any) => {
            const sections = tab.sections.get();
            sections.forEach((section: any) => {
              const sectionName = section.getName();
              const originalState = originalSectionVisibilityForSections.get(sectionName);
              if (originalState !== undefined) {
                section.setVisible(originalState);
              }
            });
          });
          // Clear the saved states after restoration
          originalSectionVisibilityForSections.clear();
        }

        result = { success: true };
        break;

      case 'GET_SCHEMA_NAMES':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }
        const attrs = Xrm.Page.data.entity.attributes.get();
        const schemaNames: string[] = [];
        attrs.forEach((attr: any) => {
          schemaNames.push(attr.getName());
        });
        result = schemaNames.sort();
        break;

      case 'TOGGLE_BLUR_FIELDS':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }

        const blurClass = 'd365-helper-blur-field';

        if (data.blur) {
          // Add blur effect to all input fields
          const inputs = document.querySelectorAll('input[type="text"], input[type="email"], input[type="url"], input[type="tel"], input[type="number"], textarea, select, [role="combobox"], [role="textbox"]');
          inputs.forEach((input: Element) => {
            (input as HTMLElement).classList.add(blurClass);
          });

          // Blur lookup values - use multiple strategies
          let lookupCount = 0;

          // Strategy 1: Use Xrm to find lookup fields and blur their displayed values
          const attributes = Xrm.Page.data.entity.attributes.get();
          attributes.forEach((attr: any) => {
            if (attr.getAttributeType() === 'lookup') {
              const controls = attr.controls.get();
              controls.forEach((control: any) => {
                try {
                  const controlName = control.getName();

                  // Try multiple selectors to find the control container
                  let controlContainer = document.querySelector(`[data-id="${controlName}"]`);
                  if (!controlContainer) {
                    controlContainer = document.querySelector(`[id="${controlName}"]`);
                  }
                  if (!controlContainer) {
                    controlContainer = document.querySelector(`[data-control-name="${controlName}"]`);
                  }

                  if (controlContainer) {
                    // Strategy A: Try to blur anchor tags first (standard lookups)
                    const lookupLinks = controlContainer.querySelectorAll('a');

                    lookupLinks.forEach((link: Element) => {
                      const linkEl = link as HTMLElement;
                      if (linkEl.textContent && linkEl.textContent.trim().length > 0) {
                        if (!linkEl.classList.contains(blurClass)) {
                          linkEl.classList.add(blurClass);
                          lookupCount++;
                        }
                      }
                    });

                    // Strategy B: If no anchor tags found, blur specific lookup value containers
                    // Look for elements with specific lookup-related attributes or classes
                    if (lookupLinks.length === 0) {
                      // Try to find elements with lookup-specific attributes
                      const lookupElements = controlContainer.querySelectorAll(
                        '[data-lp-id], [aria-label*="Lookup"], button[aria-label], [role="button"]'
                      );

                      lookupElements.forEach((el: Element) => {
                        const element = el as HTMLElement;
                        // Get the text content
                        const text = element.textContent?.trim() || '';
                        // Skip if it's just a label or empty
                        if (text.length > 0 && !text.match(/^[A-Z][a-z]+:?$/)) {
                          if (!element.classList.contains(blurClass)) {
                            element.classList.add(blurClass);
                            lookupCount++;
                          }
                        }
                      });
                    }
                  }
                } catch (e) {
                  // Silently ignore errors processing lookup controls
                }
              });
            }
          });

          // Strategy 2: Also blur any anchor tags within elements that have data-lp-id or are lookup containers
          // This catches lookups that might not be found via Xrm
          const allLookupLinks = document.querySelectorAll('[data-lp-id] a, [class*="lookup"] a, [class*="Lookup"] a');

          allLookupLinks.forEach((link: Element) => {
            const linkEl = link as HTMLElement;
            if (linkEl.textContent && linkEl.textContent.trim().length > 0 && !linkEl.classList.contains(blurClass)) {
              linkEl.classList.add(blurClass);
            }
          });
        } else {
          // Remove blur effect
          const blurredElements = document.querySelectorAll(`.${blurClass}`);
          blurredElements.forEach((element: Element) => {
            (element as HTMLElement).classList.remove(blurClass);
          });
        }

        result = { success: true };
        break;

      case 'UNLOCK_FIELDS':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }
        const allAttributes = Xrm.Page.data.entity.attributes.get();
        let unlockedCount = 0;
        allAttributes.forEach((attribute: any) => {
          const controls = attribute.controls.get();
          controls.forEach((control: any) => {
            try {
              // Try to disable the disabled state
              control.setDisabled(false);
              unlockedCount++;
            } catch (e) {
              // Some controls can't be unlocked
            }
          });
        });
        result = { unlockedCount };
        break;

      case 'AUTO_FILL_FORM':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }
        const formAttrs = Xrm.Page.data.entity.attributes.get();
        let filledCount = 0;
        formAttrs.forEach((attr: any) => {
          try {
            const attrType = attr.getAttributeType();
            const attrName = attr.getName().toLowerCase();
            const currentValue = attr.getValue();

            // Only fill if empty
            if (currentValue === null || currentValue === undefined || currentValue === '' || (Array.isArray(currentValue) && currentValue.length === 0)) {
              switch (attrType) {
                case 'string':
                case 'memo':
                  // Check if it's an email field
                  if (attrName.includes('email') || attrName.includes('emailaddress')) {
                    attr.setValue('test@example.com');
                    filledCount++;
                  }
                  // Check if it's a URL field
                  else if (attrName.includes('url') || attrName.includes('website') || attrName.includes('web')) {
                    attr.setValue('https://www.example.com');
                    filledCount++;
                  }
                  // Check if it's a phone number field
                  else if (attrName.includes('phone') || attrName.includes('telephone') || attrName.includes('mobile')) {
                    attr.setValue('555-0100');
                    filledCount++;
                  }
                  // Default text
                  else {
                    attr.setValue('Sample Text');
                    filledCount++;
                  }
                  break;
                case 'integer':
                  attr.setValue(100);
                  filledCount++;
                  break;
                case 'double':
                case 'decimal':
                case 'money':
                  attr.setValue(100.00);
                  filledCount++;
                  break;
                case 'boolean':
                  attr.setValue(true);
                  filledCount++;
                  break;
                case 'datetime':
                  attr.setValue(new Date());
                  filledCount++;
                  break;
                case 'optionset':
                  const options = attr.getOptions();
                  if (options && options.length > 0) {
                    attr.setValue(options[0].value);
                    filledCount++;
                  }
                  break;
                case 'lookup':
                  // Lookups require special handling - we can't just set a random value
                  // Skip lookups for now as they need actual entity references
                  break;
              }
            }
          } catch (e) {
            // Skip fields that can't be filled
          }
        });
        result = { filledCount };
        break;

      case 'GET_CONTROL_INFO':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }
        const allAttrs = Xrm.Page.data.entity.attributes.get();
        const controlInfo: any[] = [];
        
        // Also check header section controls if available
        let headerControls: any[] = [];
        try {
          if (Xrm.Page.ui && Xrm.Page.ui.headerSection) {
            const headerSection = Xrm.Page.ui.headerSection;
            const headerSections = headerSection.sections ? headerSection.sections.get() : [];
            headerSections.forEach((section: any) => {
              try {
                const sectionControls = section.controls.get();
                headerControls = headerControls.concat(sectionControls);
              } catch (e) {
                // Silently ignore errors getting header section controls
              }
            });
          }
        } catch (e) {
          // Silently ignore errors accessing header section
        }
        
        allAttrs.forEach((attr: any) => {
          const schemaName = attr.getName();
          const controls = attr.controls.get();
          controls.forEach((control: any) => {
            try {
              const controlName = control.getName();

              // Try multiple ways to find the element
              let element = document.getElementById(controlName);
              let elementFound = false;

              // If direct ID doesn't work, try using the control's container
              if (!element) {
                try {
                  // Try to get the control's container element directly from Xrm
                  const controlElement = control.getControlType ? control : null;
                  if (controlElement) {
                    // Look for elements with data-id matching the control name
                    const selector = `[data-id="${controlName}"], [id*="${controlName}"]`;
                    element = document.querySelector(selector);
                  }
                } catch (e) {
                  // Silently ignore errors getting control element
                }
              }

              // Additional search strategies for header fields and other cases
              if (!element) {
                // Try finding by aria-label or aria-describedby
                const ariaElements = document.querySelectorAll(`[aria-label*="${controlName}"], [aria-describedby*="${controlName}"]`);
                if (ariaElements.length > 0) {
                  element = ariaElements[0] as HTMLElement;
                }
              }

              if (!element) {
                // Try finding in header section specifically - use multiple selectors
                const headerSelectors = [
                  '[data-id="header"]',
                  '.ms-crm-Form-Header',
                  '[class*="header"]',
                  '[class*="Header"]',
                  '[id*="header"]',
                  '[id*="Header"]'
                ];
                
                for (const headerSelector of headerSelectors) {
                  const headerSection = document.querySelector(headerSelector);
                  if (headerSection) {
                    const headerControl = headerSection.querySelector(`[data-id="${controlName}"], [id*="${controlName}"], [data-lp-id="${controlName}"]`);
                    if (headerControl) {
                      element = headerControl as HTMLElement;
                      break;
                    }
                  }
                }
              }

              // If still not found, try other selectors
              if (!element) {
                const fallbackSelectors = [
                  `[data-lp-id="${controlName}"]`,
                  `[name="${controlName}"]`,
                  `input[id*="${controlName}"]`,
                  `select[id*="${controlName}"]`,
                  `textarea[id*="${controlName}"]`,
                  `[data-control-name="${controlName}"]`
                ];
                
                for (const selector of fallbackSelectors) {
                  const found = document.querySelector(selector);
                  if (found) {
                    element = found as HTMLElement;
                    break;
                  }
                }
              }

              if (element) {
                elementFound = true;
              }

              // Include controls that are visible (even if element not found yet, overlay logic will try harder)
              // For header controls, include them even if visibility check fails (they might be in header section)
              const isVisible = control.getVisible();
              const isHeaderControl = headerControls.some((hc: any) => hc.getName && hc.getName() === controlName);
              
              if (isVisible || isHeaderControl) {
                controlInfo.push({
                  schemaName: schemaName,
                  controlName: controlName,
                  label: control.getLabel ? control.getLabel() : schemaName,
                  visible: isVisible,
                  elementFound: elementFound,
                  isHeader: isHeaderControl
                });
              }
            } catch (e) {
              // Silently ignore errors processing control
            }
          });
        });
        result = controlInfo;
        break;

      case 'DISABLE_REQUIRED_FIELDS':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }
        const requiredAttributes = Xrm.Page.data.entity.attributes.get();
        let disabledCount = 0;
        requiredAttributes.forEach((attr: any) => {
          try {
            if (typeof attr.getRequiredLevel === 'function' && typeof attr.setRequiredLevel === 'function') {
              const level = attr.getRequiredLevel();
              if (level && level.toLowerCase && level.toLowerCase() !== 'none') {
                attr.setRequiredLevel('none');
                disabledCount++;
              }
            }
          } catch (e) {
            // ignore failures
          }
        });
        result = { disabledCount };
        break;

      case 'GET_OPTION_SETS':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }

        const optionAttributes = Xrm.Page.data.entity.attributes.get();
        const optionResults: any[] = [];

        optionAttributes.forEach((attr: any) => {
          try {
            if (!attr || typeof attr.getAttributeType !== 'function') {
              return;
            }

            const attributeType = attr.getAttributeType && attr.getAttributeType();
            if (!['optionset', 'multioptionset', 'boolean'].includes(attributeType)) {
              return;
            }

            const logicalName = typeof attr.getName === 'function' ? attr.getName() : 'unknown';
            const controls = attr.controls && attr.controls.get ? attr.controls.get() : [];
            let displayLabel = logicalName;
            if (controls && controls.length > 0 && typeof controls[0].getLabel === 'function') {
              displayLabel = controls[0].getLabel() || displayLabel;
            }

            let options: any[] = [];
            if (typeof attr.getOptions === 'function') {
              options = attr.getOptions() || [];
            }

            const mappedOptions = options.map((option: any) => {
              const label = option?.text || option?.label?.LocalizedLabels?.[0]?.Label || option?.label?.UserLocalizedLabel?.Label || (option?.value != null ? option.value.toString() : 'Unnamed');
              return {
                value: option?.value,
                label,
                color: option?.color || option?.Color || null,
                isDefault: Boolean(option?.defaultSelected || option?.Default || option?.isDefault)
              };
            });

            if (mappedOptions.length === 0) {
              return;
            }

            const currentValue = typeof attr.getValue === 'function' ? attr.getValue() : null;
            const currentValueLabel = typeof attr.getText === 'function' ? attr.getText() : '';
            const isMultiSelect = attributeType === 'multioptionset';

            optionResults.push({
              logicalName,
              displayLabel,
              attributeType,
              isMultiSelect,
              currentValue,
              currentValueLabel,
              optionCount: mappedOptions.length,
              options: mappedOptions
            });
          } catch (e) {
          }
        });

        optionResults.sort((a, b) => (a.displayLabel || '').localeCompare(b.displayLabel || ''));

        result = {
          attributes: optionResults
        };
        break;

      case 'GET_FORM_LIBRARIES':
        const libraries: any[] = [];
        const onLoadHandlers: any[] = [];
        const onChangeHandlers: any[] = [];
        const onSaveHandlers: any[] = [];

        // Check if we have access to form data
        if (!Xrm || !Xrm.Page) {
          console.warn('D365 Helper: Xrm.Page not available');
          result = {
            libraries: [],
            onLoad: [],
            onChange: [],
            onSave: [],
            error: 'Xrm API not available. Please wait for the page to fully load.'
          };
          break;
        }

        if (!Xrm.Page.data || !Xrm.Page.data.entity) {
          // Detect if we're on a list view
          const url = window.location.href;
          const isListView = url.includes('pagetype=entitylist') || url.includes('viewid=');

          result = {
            libraries: [],
            onLoad: [],
            onChange: [],
            onSave: [],
            error: isListView
              ? 'You are on a list view. Please open a record to view its JavaScript libraries and event handlers.'
              : 'Form data not available. This feature works on entity form pages only.'
          };
          break;
        }

        // Method 0: Try official Xrm API to get registered event handlers
        try {
          // Check if form-level event methods exist
          const entity = Xrm.Page.data.entity as any;

          // Try to get OnLoad handlers - there's no official "get" method but we can try reflection

          // OnChange for attributes using official API
          const attributes = Xrm.Page.data.entity.attributes.get();
          attributes.forEach((attr: any) => {
            try {
              // Some D365 versions have addOnChange with registered handlers
              if (typeof attr.addOnChange === 'function') {
                // The function itself might have a reference to registered handlers
                const onChangeFn = attr.addOnChange;
                if ((onChangeFn as any)._handlers) {
                }
              }

              // Try to call a fake handler to see what's registered (won't execute)
              // This is a hack but might reveal registered handlers in error messages
            } catch (e) {
              // Silent
            }
          });

          // Try accessing form-level OnLoad
          if (typeof entity.addOnLoad === 'function') {
            const addOnLoad = entity.addOnLoad;
            if ((addOnLoad as any)._handlers) {
            }
          }

          // Try accessing OnSave
          if (typeof entity.addOnSave === 'function') {
            const addOnSave = entity.addOnSave;
            if ((addOnSave as any)._handlers) {
            }
          }
        } catch (e) {
        }

        // Try multiple approaches to get event handlers

        // Method 1: Check _clientApiExecutor._store (new D365 approach)
        try {
          const entity = Xrm.Page.data.entity as any;
          const executor = entity._clientApiExecutor;

          if (executor) {

            // Check the _store property - this is a Redux store
            if (executor._store) {

              const store = executor._store;

              // Get state from Redux store
              if (typeof store.getState === 'function') {
                try {
                  const state = store.getState();

                  // Look for form libraries and event handlers in state
                  if (state.formLibraries) {
                  }

                  if (state.libraries) {
                  }

                  if (state.eventHandlers) {
                  }

                  if (state.events) {
                  }

                  if (state.handlers) {
                  }

                  // Check for form metadata
                  if (state.form) {
                    if (state.form.libraries) {
                    }
                  }

                  if (state.formData) {
                  }

                  // Check pages state (likely contains form metadata)
                  if (state.pages) {

                    const pages = state.pages;
                    // Pages might be keyed by page ID
                    const pageKeys = Object.keys(pages);
                    if (pageKeys.length > 0) {
                      const firstPage = pages[pageKeys[0]];
                      if (firstPage) {

                        // Look for form libraries and event handlers
                        if (firstPage.formLibraries) {
                        }
                        if (firstPage.libraries) {
                        }
                        if (firstPage.eventHandlers) {
                        }
                        if (firstPage.events) {
                        }
                        if (firstPage.handlers) {
                        }
                        if (firstPage.controls) {
                        }
                        if (firstPage.data) {
                        }

                        // Check forms in page
                        if (firstPage.forms) {
                          const formKeys = Object.keys(firstPage.forms);
                          if (formKeys.length > 0) {
                            const firstForm = firstPage.forms[formKeys[0]];

                            if (firstForm.formLibraries) {
                            }
                            if (firstForm.events) {
                            }
                          }
                        }

                        // Check metadata in page
                        if (firstPage.metadata) {
                        }
                      }
                    }
                  }

                  // Check metadata state (might have form definitions)
                  if (state.metadata) {

                    const metadata = state.metadata;

                    // Check forms metadata
                    if (metadata.forms) {

                      // Get current form ID
                      const currentFormId = (Xrm.Page.ui.formSelector?.getCurrentItem()?.getId() || '').replace(/[{}]/g, '');

                      if (currentFormId && metadata.forms[currentFormId]) {
                        const formMetadata = metadata.forms[currentFormId];

                        // Extract form libraries (check both PascalCase and camelCase)
                        const formLibraries = formMetadata.FormLibraries || formMetadata.formLibraries;
                        if (formLibraries) {

                          if (Array.isArray(formLibraries)) {
                            formLibraries.forEach((lib: any) => {
                              const libName = lib.Name || lib.name || lib.LibraryName || lib.libraryName || lib;
                              if (libName && typeof libName === 'string' && !libraries.find(l => l.name === libName)) {
                                libraries.push({ name: libName, order: lib.Order || lib.order || 0 });
                              }
                            });
                          }
                        }

                        // Extract event handlers - EventHandlers is an ARRAY of handler objects
                        const eventHandlers = formMetadata.EventHandlers || formMetadata.eventHandlers || formMetadata.events;
                        if (eventHandlers && Array.isArray(eventHandlers)) {

                          eventHandlers.forEach((handler: any) => {
                            const eventName = (handler.EventName || handler.eventName || '').toLowerCase();
                            const functionName = handler.FunctionName || handler.functionName || handler.name || handler.Name;
                            const libraryName = handler.LibraryName || handler.libraryName || handler.library || handler.Library;
                            const attributeName = handler.AttributeName || handler.attributeName;
                            const enabled = handler.Enabled !== false && handler.enabled !== false;

                            if (!functionName) return;

                            // OnLoad handlers (no AttributeName)
                            if (eventName === 'onload' && !attributeName) {
                              onLoadHandlers.push({
                                type: 'form',
                                target: 'Form',
                                library: libraryName || 'Unknown',
                                functionName: functionName,
                                enabled: enabled
                              });
                            }
                            // OnChange handlers (has AttributeName)
                            else if (eventName === 'onchange' && attributeName) {
                              onChangeHandlers.push({
                                type: 'field',
                                target: attributeName,
                                library: libraryName || 'Unknown',
                                functionName: functionName,
                                enabled: enabled
                              });
                            }
                            // OnSave handlers (no AttributeName)
                            else if (eventName === 'onsave' && !attributeName) {
                              onSaveHandlers.push({
                                type: 'form',
                                target: 'Form',
                                library: libraryName || 'Unknown',
                                functionName: functionName,
                                enabled: enabled
                              });
                            }
                          });
                        }

                        // Check for controls with onChange events (use Controls with PascalCase)
                        const controls = formMetadata.Controls || formMetadata.controls;
                        if (controls) {

                          Object.keys(controls).forEach((controlKey) => {
                            const control = controls[controlKey];
                            const controlEvents = control.EventHandlers || control.eventHandlers || control.events || control.Events;

                            if (controlEvents) {
                              if (controlEvents.onchange || controlEvents.onChange || controlEvents.OnChange) {
                                const onChangeEvents = controlEvents.onchange || controlEvents.onChange || controlEvents.OnChange;

                                if (Array.isArray(onChangeEvents)) {
                                  onChangeEvents.forEach((handler: any) => {
                                    const functionName = handler.FunctionName || handler.functionName || handler.name || handler.Name;
                                    const libraryName = handler.LibraryName || handler.libraryName || handler.library || handler.Library;
                                    const fieldName = control.DataFieldName || control.datafieldname || control.Id || control.id || controlKey;

                                    if (functionName) {
                                      onChangeHandlers.push({
                                        type: 'field',
                                        target: fieldName,
                                        library: libraryName || 'Unknown',
                                        functionName: functionName,
                                        enabled: handler.Enabled !== false && handler.enabled !== false
                                      });
                                    }
                                  });
                                }
                              }
                            }
                          });
                        }
                      }
                    }
                  }

                  // Log all state keys for debugging
                } catch (stateError) {
                }
              }

              // Look for event handlers in store
              if (store._eventHandlers) {
              }

              if (store.eventHandlers) {
              }

              // Check for onLoad handlers
              if (store.onload || store.onLoad || store.OnLoad) {
                const handlers = store.onload || store.onLoad || store.OnLoad;
                if (Array.isArray(handlers)) {
                  handlers.forEach((handler: any) => {
                    if (handler && (handler.functionName || handler.name)) {
                      onLoadHandlers.push({
                        type: 'form',
                        target: 'Form',
                        library: handler.libraryName || handler.library || 'Unknown',
                        functionName: handler.functionName || handler.name,
                        enabled: handler.enabled !== false
                      });
                    }
                  });
                }
              }

              // Check for onSave handlers
              if (store.onsave || store.onSave || store.OnSave) {
                const handlers = store.onsave || store.onSave || store.OnSave;
                if (Array.isArray(handlers)) {
                  handlers.forEach((handler: any) => {
                    if (handler && (handler.functionName || handler.name)) {
                      onSaveHandlers.push({
                        type: 'form',
                        target: 'Form',
                        library: handler.libraryName || handler.library || 'Unknown',
                        functionName: handler.functionName || handler.name,
                        enabled: handler.enabled !== false
                      });
                    }
                  });
                }
              }

              // Try to find libraries in store
              if (store.libraries || store.formLibraries) {
                const libs = store.libraries || store.formLibraries;
                if (Array.isArray(libs)) {
                  libs.forEach((lib: any) => {
                    const libName = lib.name || lib.libraryName || lib;
                    if (libName && typeof libName === 'string') {
                      libraries.push({ name: libName, order: lib.order || 0 });
                    }
                  });
                }
              }
            }

            // Try to find event handlers in the executor itself
            if (executor._eventHandlers) {
            }

            // Check for registered events
            if (executor._registeredEvents) {
            }
          }
        } catch (e) {
        }

        // Method 2: Check _formContext (might contain event info)
        try {
          const data = Xrm.Page.data as any;
          const formContext = data._formContext;

          if (formContext) {

            // Deep inspect formContext
            if (formContext._eventHandlers) {
            }

            if (formContext.data) {
              if ((formContext.data as any)._eventHandlers) {
              }
            }
          }
        } catch (e) {
        }

        // Method 3: Check legacy _eventHandlers property
        try {
          const entity = Xrm.Page.data.entity as any;

          if (entity._eventHandlers) {
            // OnLoad
            if (entity._eventHandlers.onload) {
              entity._eventHandlers.onload.forEach((handler: any) => {
                if (handler && handler.functionName) {
                  onLoadHandlers.push({
                    type: 'form',
                    target: 'Form',
                    library: handler.libraryName || 'Unknown',
                    functionName: handler.functionName,
                    enabled: handler.enabled !== false
                  });
                }
              });
            }

            // OnSave
            if (entity._eventHandlers.onsave) {
              entity._eventHandlers.onsave.forEach((handler: any) => {
                if (handler && handler.functionName) {
                  onSaveHandlers.push({
                    type: 'form',
                    target: 'Form',
                    library: handler.libraryName || 'Unknown',
                    functionName: handler.functionName,
                    enabled: handler.enabled !== false
                  });
                }
              });
            }
          }
        } catch (e) {
        }

        // Method 4: Check attributes for onChange handlers via _clientApiExecutorAttribute
        try {
          const formAttributes = Xrm.Page.data.entity.attributes.get();

          // Just check first attribute in detail to avoid too much logging
          if (formAttributes.length > 0) {
            const firstAttr = formAttributes[0] as any;
            const fieldName = firstAttr.getName();

            if (firstAttr._clientApiExecutorAttribute) {

              const attrExecutor = firstAttr._clientApiExecutorAttribute;

              // Check for _store in attribute executor
              if (attrExecutor._store) {

                const attrStore = attrExecutor._store;

                // Get state from attribute's Redux store
                if (typeof attrStore.getState === 'function') {
                  try {
                    const attrState = attrStore.getState();

                    if (attrState.eventHandlers || attrState.events || attrState.onchange) {
                    }
                  } catch (e) {
                  }
                }

                if (attrStore.onchange || attrStore.onChange || attrStore.OnChange) {
                }
              }

              if (attrExecutor._eventHandlers) {
              }
            }
          }

          // Now check all attributes for onChange handlers
          formAttributes.forEach((attr: any) => {
            try {
              const fieldName = attr.getName();

              // Check _eventHandlers on attribute
              if (attr._eventHandlers?.onchange) {
                attr._eventHandlers.onchange.forEach((handler: any) => {
                  if (handler && (handler.functionName || handler.name)) {
                    onChangeHandlers.push({
                      type: 'field',
                      target: fieldName,
                      library: handler.libraryName || handler.library || 'Unknown',
                      functionName: handler.functionName || handler.name,
                      enabled: handler.enabled !== false
                    });
                  }
                });
              }

              // Check _clientApiExecutorAttribute._store.onchange
              const attrExecutor = attr._clientApiExecutorAttribute;
              if (attrExecutor?._store) {
                const attrStore = attrExecutor._store;
                const changeHandlers = attrStore.onchange || attrStore.onChange || attrStore.OnChange;

                if (changeHandlers && Array.isArray(changeHandlers)) {
                  changeHandlers.forEach((handler: any) => {
                    if (handler && (handler.functionName || handler.name)) {
                      onChangeHandlers.push({
                        type: 'field',
                        target: fieldName,
                        library: handler.libraryName || handler.library || 'Unknown',
                        functionName: handler.functionName || handler.name,
                        enabled: handler.enabled !== false
                      });
                    }
                  });
                }
              }

              // Check _clientApiExecutorAttribute._eventHandlers
              if (attrExecutor?._eventHandlers?.onchange) {
                attrExecutor._eventHandlers.onchange.forEach((handler: any) => {
                  if (handler && (handler.functionName || handler.name)) {
                    onChangeHandlers.push({
                      type: 'field',
                      target: fieldName,
                      library: handler.libraryName || handler.library || 'Unknown',
                      functionName: handler.functionName || handler.name,
                      enabled: handler.enabled !== false
                    });
                  }
                });
              }
            } catch (e) {
            }
          });
        } catch (e) {
        }

        // Method 5: Access form XML metadata (where event handlers are defined)
        try {
          // Try to access form XML from various places
          const data = Xrm.Page.data as any;
          const entity = Xrm.Page.data.entity as any;

          // Check _xrmForm which might contain form XML
          if (entity._xrmForm) {

            const xrmForm = entity._xrmForm;

            // Look for form XML or form descriptor
            if (xrmForm.FormXml || xrmForm.formXml) {
              const formXml = xrmForm.FormXml || xrmForm.formXml;

              if (formXml && typeof formXml === 'string') {
                // Parse the XML to extract libraries and event handlers
                try {
                  const parser = new DOMParser();
                  const xmlDoc = parser.parseFromString(formXml, 'text/xml');

                  // Extract form libraries
                  const formLibraries = xmlDoc.querySelectorAll('Library');

                  formLibraries.forEach((lib: any) => {
                    const libName = lib.getAttribute('name');
                    const libOrder = parseInt(lib.getAttribute('order') || '999');
                    if (libName) {
                      if (!libraries.find(l => l.name === libName)) {
                        libraries.push({ name: libName, order: libOrder });
                      }
                    }
                  });

                  // Extract OnLoad event handlers
                  const onLoadEvents = xmlDoc.querySelectorAll('event[name="onload"] Handler, event[name="OnLoad"] Handler');

                  onLoadEvents.forEach((handler: any) => {
                    const functionName = handler.getAttribute('functionName');
                    const libraryName = handler.getAttribute('libraryName');
                    const enabled = handler.getAttribute('enabled') !== 'false';

                    if (functionName) {
                      onLoadHandlers.push({
                        type: 'form',
                        target: 'Form',
                        library: libraryName || 'Unknown',
                        functionName: functionName,
                        enabled: enabled
                      });
                    }
                  });

                  // Extract OnSave event handlers
                  const onSaveEvents = xmlDoc.querySelectorAll('event[name="onsave"] Handler, event[name="OnSave"] Handler');

                  onSaveEvents.forEach((handler: any) => {
                    const functionName = handler.getAttribute('functionName');
                    const libraryName = handler.getAttribute('libraryName');
                    const enabled = handler.getAttribute('enabled') !== 'false';

                    if (functionName) {
                      onSaveHandlers.push({
                        type: 'form',
                        target: 'Form',
                        library: libraryName || 'Unknown',
                        functionName: functionName,
                        enabled: enabled
                      });
                    }
                  });

                  // Extract OnChange event handlers for fields
                  const onChangeEvents = xmlDoc.querySelectorAll('control event[name="onchange"] Handler, control event[name="OnChange"] Handler');

                  onChangeEvents.forEach((handler: any) => {
                    const functionName = handler.getAttribute('functionName');
                    const libraryName = handler.getAttribute('libraryName');
                    const enabled = handler.getAttribute('enabled') !== 'false';

                    // Get the parent control to find field name
                    const control = handler.closest('control');
                    const fieldName = control?.getAttribute('datafieldname') || control?.getAttribute('id') || 'Unknown';

                    if (functionName) {
                      onChangeHandlers.push({
                        type: 'field',
                        target: fieldName,
                        library: libraryName || 'Unknown',
                        functionName: functionName,
                        enabled: enabled
                      });
                    }
                  });
                } catch (xmlError) {
                }
              }
            }

            // Check for FormDescriptor
            if (xrmForm.FormDescriptor || xrmForm.formDescriptor) {
              const desc = xrmForm.FormDescriptor || xrmForm.formDescriptor;
            }
          }

          // Try alternate locations for form XML
          if (data._formContext) {
            const formContext = data._formContext;

            if ((formContext as any).FormXml) {
            }
          }
        } catch (e) {
        }

        // Method 6: Scan for scripts in DOM and check window object for library namespaces
        try {
          // Look for WebResources (case-insensitive)
          const allScripts = document.querySelectorAll('script[src]');

          const webResourceScripts: any[] = [];
          const customLibNames: string[] = [];

          allScripts.forEach((script: any) => {
            const src = script.getAttribute('src') || '';
            if (src.toLowerCase().includes('webresource')) {
              webResourceScripts.push(script);

              // Extract library name
              const jsMatch = src.match(/([^\/]+\.js)(?:\?|$)/i);
              if (jsMatch) {
                const libName = jsMatch[1];

                // Filter out system libraries
                if (!libName.includes('SaveWebresourcesVersions') &&
                    !libName.includes('rtejsanity') &&
                    !libName.includes('fluentui_react')) {
                  customLibNames.push(libName);
                }
              }
            }
          });


          // Add custom libraries to the list
          webResourceScripts.forEach((script: any) => {
            const src = script.getAttribute('src');
            if (src) {
              let libName = '';
              const jsMatch = src.match(/([^\/]+\.js)(?:\?|$)/i);
              if (jsMatch) {
                libName = jsMatch[1];
              }

              if (libName && !libraries.find(lib => lib.name === libName)) {
                libraries.push({ name: libName, order: 999 });
              }
            }
          });

          // Try to find library namespaces in window object
          // Common patterns: window.MyNamespace, window.CompanyName, etc.
          const windowKeys = Object.keys(window);

          // Look for non-standard properties that might be custom libraries
          const suspectedNamespaces = windowKeys.filter(key => {
            // Skip known browser/D365 globals
            const skipList = ['Xrm', 'parent', 'opener', 'top', 'window', 'self', 'frames',
                            'document', 'location', 'navigator', 'screen', 'history',
                            'localStorage', 'sessionStorage', 'console', 'jQuery', '$',
                            'Microsoft', 'ClientGlobalContext'];

            if (skipList.includes(key)) return false;
            if (key.startsWith('_') || key.startsWith('webkit')) return false;
            if (typeof (window as any)[key] !== 'object' || (window as any)[key] === null) return false;

            // Check if it's an object with functions (likely a namespace)
            const obj = (window as any)[key];
            if (typeof obj === 'object' && obj !== null) {
              const objKeys = Object.keys(obj);
              const hasFunctions = objKeys.some(k => typeof obj[k] === 'function');
              return hasFunctions && objKeys.length > 0 && objKeys.length < 100;
            }

            return false;
          });

          if (suspectedNamespaces.length > 0) {
            suspectedNamespaces.forEach(ns => {
              const obj = (window as any)[ns];
              const functions = Object.keys(obj).filter(k => typeof obj[k] === 'function');
            });
          }
        } catch (e) {
        }

        // Deduplicate library names from handlers
        const librarySet = new Set<string>();
        [...onLoadHandlers, ...onChangeHandlers, ...onSaveHandlers].forEach((handler: any) => {
          if (handler.library && handler.library !== 'Unknown') {
            librarySet.add(handler.library);
          }
        });

        // Add libraries found in handlers if not already in libraries list
        librarySet.forEach((libName: string) => {
          if (!libraries.find(lib => lib.name === libName)) {
            libraries.push({ name: libName, order: 999 });
          }
        });


        result = {
          libraries: libraries.sort((a, b) => a.order - b.order),
          onLoad: onLoadHandlers,
          onChange: onChangeHandlers,
          onSave: onSaveHandlers
        };
        break;

      case 'GET_ODATA_FIELDS':
        if (!Xrm || !Xrm.Page || !Xrm.Page.data || !Xrm.Page.data.entity) {
          throw new Error('Form not fully loaded. Please wait for the form to finish loading.');
        }

        try {

          // Get entity name and client URL
          const entityLogicalName = Xrm.Page.data.entity.getEntityName();
          const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();


          // Step 1: Fetch basic entity metadata
          const entityMetadataUrl = `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')?$select=EntitySetName,SchemaName,LogicalName`;


          const entityResponse = await fetch(entityMetadataUrl, {
            method: 'GET',
            headers: {
              'Accept': 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0'
            },
            credentials: 'include'
          });


          if (!entityResponse.ok) {
            const errorText = await entityResponse.text();
            console.error('D365 Helper: Entity metadata error:', errorText);
            throw new Error(`Failed to fetch entity metadata: ${entityResponse.statusText}`);
          }

          const entityMetadataResult = await entityResponse.json();
          const entitySetName = entityMetadataResult.EntitySetName;
          const entitySchemaName = entityMetadataResult.SchemaName;


          // Step 2: Fetch attributes separately
          const attributesUrl = `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes?$select=LogicalName,SchemaName,AttributeType`;

          const attributesResponse = await fetch(attributesUrl, {
            method: 'GET',
            headers: {
              'Accept': 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0'
            },
            credentials: 'include'
          });

          if (!attributesResponse.ok) {
            const errorText = await attributesResponse.text();
            console.error('D365 Helper: Attributes error:', errorText);
            throw new Error(`Failed to fetch attributes: ${attributesResponse.statusText}`);
          }

          const attributesData = await attributesResponse.json();
          const attributes = attributesData.value || [];


          // Step 3: Fetch relationship metadata for lookups
          const relationshipsUrl = `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')/ManyToOneRelationships?$select=SchemaName,ReferencedEntity,ReferencingEntity,ReferencingAttribute`;

          const relationshipsResponse = await fetch(relationshipsUrl, {
            method: 'GET',
            headers: {
              'Accept': 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0'
            },
            credentials: 'include'
          });

          let relationshipsMap = new Map<string, any>();
          if (relationshipsResponse.ok) {
            const relationshipsData = await relationshipsResponse.json();
            const relationships = relationshipsData.value || [];


            // Map relationships by ReferencingAttribute (which is the lookup field logical name)
            relationships.forEach((rel: any) => {
              if (rel.ReferencingAttribute) {
                relationshipsMap.set(rel.ReferencingAttribute, {
                  schemaName: rel.SchemaName,
                  referencedEntity: rel.ReferencedEntity,
                  referencingEntity: rel.ReferencingEntity
                });
              }
            });
          } else {
          }

          // Create a map of attributes from the form to get option set values and targets
          const formAttributesMap = new Map<string, any>();
          const formAttributes = Xrm.Page.data.entity.attributes.get();
          formAttributes.forEach((attr: any) => {
            try {
              const logicalName = attr.getName();
              const attributeType = attr.getAttributeType();

              const attrData: any = {
                type: attributeType
              };

              // Get option set values for picklists
              if (['optionset', 'multioptionset', 'boolean'].includes(attributeType) && typeof attr.getOptions === 'function') {
                const options = attr.getOptions() || [];
                attrData.options = options.map((opt: any) => ({
                  value: opt.value,
                  label: opt.text || opt.label || ''
                }));
              }

              formAttributesMap.set(logicalName, attrData);
            } catch (e) {
              // Skip attributes that can't be accessed
            }
          });

          // Now fetch detailed metadata for lookups and option sets
          // We'll make individual calls for each attribute type we're interested in
          const lookupFieldsUrl = `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')/Attributes/Microsoft.Dynamics.CRM.LookupAttributeMetadata?$select=LogicalName,Targets`;
          let lookupTargets = new Map<string, string[]>();

          try {
            const lookupResponse = await fetch(lookupFieldsUrl, {
              method: 'GET',
              headers: {
                'Accept': 'application/json',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0'
              },
              credentials: 'include'
            });

            if (lookupResponse.ok) {
              const lookupData = await lookupResponse.json();
              (lookupData.value || []).forEach((lookup: any) => {
                if (lookup.Targets && lookup.Targets.length > 0) {
                  lookupTargets.set(lookup.LogicalName, lookup.Targets);
                }
              });
            }
          } catch (e) {
          }

          // Process attributes to extract OData-relevant information
          const odataFields: any[] = [];

          attributes.forEach((attr: any) => {
            const field: any = {
              logicalName: attr.LogicalName,
              schemaName: attr.SchemaName,
              attributeType: attr.AttributeType
            };

            // Add OData bind for lookups
            if (attr.AttributeType === 'Lookup' || attr.AttributeType === 'Customer' || attr.AttributeType === 'Owner') {
              field.odataBind = '@odata.bind';

              // Get relationship information for this lookup
              const relationship = relationshipsMap.get(attr.LogicalName);
              if (relationship) {
                field.relationshipName = relationship.schemaName;
                field.targetEntity = relationship.referencedEntity;
              } else {
                // Fallback: try to get target from our lookup targets map
                const targets = lookupTargets.get(attr.LogicalName);
                if (targets && targets.length > 0) {
                  field.targetEntity = targets.join(', ');
                }
              }
            }

            // Get option set values from form attributes (more reliable than API)
            const formAttr = formAttributesMap.get(attr.LogicalName);
            if (formAttr && formAttr.options && formAttr.options.length > 0) {
              field.optionSetValues = formAttr.options
                .map((opt: any) => `${opt.value}-${opt.label}`)
                .join(', ');
            }

            odataFields.push(field);
          });

          // Sort fields by logical name
          odataFields.sort((a, b) => a.logicalName.localeCompare(b.logicalName));

          result = {
            entityName: entityLogicalName,
            entitySchemaName: entitySchemaName,
            entitySetName: entitySetName,
            fields: odataFields
          };
        } catch (error: any) {
          console.error('D365 Helper: Error fetching OData fields:', error);
          throw new Error(`Failed to fetch OData fields: ${error.message}`);
        }
        break;

      case 'GET_PLUGIN_TRACE_LOGS':
        if (!Xrm || !Xrm.WebApi || typeof Xrm.WebApi.retrieveMultipleRecords !== 'function') {
          result = {
            logs: [],
            error: 'Xrm.WebApi not available. Make sure you are in the Dynamics 365 app.'
          };
          break;
        }

        try {
          const top = Math.min(Math.max(Number(data?.top) || 20, 1), 200);
          const query = `?$orderby=createdon desc&$top=${top}`;

          // Try multiple entity names as D365 versions differ
          let response;
          let entityNameUsed = '';
          const entityNames = ['plugintracelog', 'plugintracelogbase', 'plugintypetracelog'];

          for (const entityName of entityNames) {
            try {
              response = await Xrm.WebApi.retrieveMultipleRecords(entityName, query);
              entityNameUsed = entityName;
              break;
            } catch (err: any) {
              // Continue to next entity name
            }
          }

          if (!response) {
            throw new Error('Unable to retrieve plugin trace logs. Tried entity names: ' + entityNames.join(', '));
          }

          const logs = response.entities.map((log: any) => ({
            id:
              log.plugintracelogid ||
              log.plugintypetracelogid ||
              log['plugintypetracelogid'] ||
              log.plugintypetracelogid_guid ||
              Math.random().toString(36).slice(2),
            createdOn: log.createdon,
            messageName: log.messagename,
            primaryEntity: log.primaryentity,
            typeName: log.typename,
            mode: log.mode,
            depth: log.depth,
            operationCorrelationId:
              log.operationcorrelationid ||
              log.correlationid ||
              log.operationCorrelationId,
            performanceDurationMs:
              log.performanceexecutionduration ??
              log.executionduration ??
              log.processduration,
            executionStart:
              log.performanceexecutionstarttime ||
              log.executionstarttime ||
              log.PerformanceExecutionStartTime,
            requestId: log.requestid || log.RequestId,
            exceptionDetails: log.exceptiondetails || log.ExceptionDetails,
            messageBlock: log.messageblock || log.messagelog || log.MessageBlock || log.MessageLog,
            createdByName: log['_createdby_value@OData.Community.Display.V1.FormattedValue'] || '',
            createdById: log._createdby_value || log['_createdby_value'] || ''
          }));

          result = {
            logs,
            moreRecords: Boolean(response.nextLink)
          };
        } catch (error: any) {
          console.error('D365 Helper: Failed to retrieve plugin trace logs', error);
          const status = error?.status ?? error?.httpStatusCode;
          const rawMessage = typeof error?.message === 'string' ? error.message : '';
          let errorMessage =
            'Failed to retrieve plugin trace logs. Ensure tracing is enabled and you have permission to read the Plug-in Trace Log table.';

          if (status === 404) {
            errorMessage =
              'The Plug-in Trace Log table is not available. Enable plug-in trace logging (Settings → Administration → System Settings → Customization) and confirm your solution includes the Plug-in Trace Log table.';
          } else if (status === 403 || status === 401) {
            errorMessage =
              'You do not have permission to read plug-in trace logs. Ask your administrator for Read access to the Plug-in Trace Log table.';
          } else if (rawMessage) {
            errorMessage = rawMessage;
          }

          result = {
            logs: [],
            error: errorMessage
          };
        }
        break;

      case 'GET_AUDIT_HISTORY':
        if (!Xrm || !Xrm.WebApi || typeof Xrm.WebApi.retrieveMultipleRecords !== 'function') {
          result = {
            records: [],
            error: 'Xrm.WebApi not available. Make sure you are in the Dynamics 365 app.'
          };
          break;
        }

        try {
          // Get current record ID and entity name
          const recordId = Xrm.Page.data.entity.getId().replace(/[{}]/g, '');
          const entityName = Xrm.Page.data.entity.getEntityName();
          const recordName = Xrm.Page.getAttribute('name')?.getValue() ||
                           Xrm.Page.data.entity.getPrimaryAttributeValue() ||
                           'Current Record';


          // Hoisted here so the audit fetch can reuse it (the existing helper-init block
          // below also redeclares this for backwards-compat reasons; keep both).
          const auditClientUrl = Xrm.Utility.getGlobalContext().getClientUrl();

          // Query audit records for this specific record. Use raw fetch so we can request
          // FormattedValue annotations (saves an N+1 user lookup) and follow @odata.nextLink
          // for records with more than one page of audit history.
          const MAX_AUDIT_RECORDS = 1000;
          const PAGE_SIZE = 250;

          const fetchAuditPage = async (url: string) => {
            const resp = await fetch(url, {
              method: 'GET',
              headers: {
                Accept: 'application/json',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                Prefer: `odata.include-annotations="*",odata.maxpagesize=${PAGE_SIZE}`,
              },
              credentials: 'include',
            });
            if (!resp.ok) {
              const text = await resp.text().catch(() => '');
              throw new Error(`HTTP ${resp.status}: ${text.slice(0, 200) || resp.statusText}`);
            }
            return resp.json();
          };

          const auditEntities: any[] = [];
          let auditTruncated = false;
          let auditUrl =
            `${auditClientUrl}/api/data/v9.2/audits` +
            `?$filter=_objectid_value eq ${recordId}` +
            `&$orderby=createdon desc`;

          // Paginate via @odata.nextLink up to MAX_AUDIT_RECORDS.
          while (auditUrl) {
            const page = await fetchAuditPage(auditUrl);
            const value: any[] = Array.isArray(page?.value) ? page.value : [];
            auditEntities.push(...value);
            const nextLink: string | undefined = page?.['@odata.nextLink'];
            if (!nextLink) {
              auditUrl = '';
              break;
            }
            if (auditEntities.length >= MAX_AUDIT_RECORDS) {
              auditTruncated = true;
              auditUrl = '';
              break;
            }
            auditUrl = nextLink;
          }

          const auditRecords: any[] = [];
          // Cache user names so 200 audits by the same user only cost one fetch
          // (and most aren't needed thanks to the FormattedValue annotation we now request).
          const userNameCache = new Map<string, string>();
          const resolveUserName = async (userId: string, fallback?: string): Promise<string> => {
            if (fallback && fallback.trim()) return fallback;
            if (!userId) return 'Unknown User';
            if (userNameCache.has(userId)) return userNameCache.get(userId) || 'Unknown User';
            try {
              const user = await Xrm.WebApi.retrieveRecord('systemuser', userId, '?$select=fullname');
              const name = user?.fullname || 'Unknown User';
              userNameCache.set(userId, name);
              return name;
            } catch {
              userNameCache.set(userId, 'Unknown User');
              return 'Unknown User';
            }
          };

          const optionLabelsByField = new Map<string, Map<number, string>>();
          const attributeTypeByField = new Map<string, string>();
          const lookupTargetsByField = new Map<string, string[]>();
          const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();

          // Fetch attribute type metadata for the whole entity so lookups not on the
          // current form are still detected (and their values resolved) in audit history.
          try {
            const escapeForUrl = (value: string): string => value.replace(/'/g, "''");
            const attrMetaUrl =
              `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${escapeForUrl(entityName)}')` +
              `/Attributes?$select=LogicalName,AttributeType`;
            const attrMetaResponse = await fetch(attrMetaUrl, {
              method: 'GET',
              headers: {
                Accept: 'application/json',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
              },
              credentials: 'include',
            });
            if (attrMetaResponse.ok) {
              const attrMetaJson = await attrMetaResponse.json();
              const rows: any[] = Array.isArray(attrMetaJson?.value) ? attrMetaJson.value : [];
              rows.forEach((row) => {
                const logical = typeof row?.LogicalName === 'string' ? row.LogicalName.toLowerCase() : '';
                const type = typeof row?.AttributeType === 'string' ? row.AttributeType.toLowerCase() : '';
                if (!logical || !type) return;
                if (!attributeTypeByField.has(logical)) {
                  attributeTypeByField.set(logical, type);
                }
              });
            }
          } catch {
            // Non-fatal: form-derived metadata below still covers visible fields.
          }

          try {
            const formAttributes = Xrm.Page.data.entity.attributes.get();
            formAttributes.forEach((attr: any) => {
              try {
                const logicalName = typeof attr.getName === 'function' ? attr.getName() : '';
                const attrType = typeof attr.getAttributeType === 'function' ? attr.getAttributeType() : '';
                if (!logicalName) return;
                const fieldKey = logicalName.toLowerCase();

                if (attrType) {
                  attributeTypeByField.set(fieldKey, String(attrType).toLowerCase());
                }

                if (['lookup', 'customer', 'owner'].includes(String(attrType).toLowerCase())) {
                  const lookupTypes =
                    typeof attr.getLookupTypes === 'function' ? attr.getLookupTypes() : [];
                  if (Array.isArray(lookupTypes) && lookupTypes.length > 0) {
                    lookupTargetsByField.set(
                      fieldKey,
                      lookupTypes
                        .map((target: any) => String(target || '').trim())
                        .filter((target: string) => target.length > 0)
                    );
                  }
                }

                if (!['optionset', 'boolean', 'multioptionset'].includes(attrType)) {
                  return;
                }

                const options = typeof attr.getOptions === 'function' ? attr.getOptions() : [];
                if (!Array.isArray(options) || options.length === 0) return;

                const labels = new Map<number, string>();
                options.forEach((option: any) => {
                  const value = Number(option?.value);
                  const label = option?.text || option?.label || '';
                  if (Number.isFinite(value) && typeof label === 'string' && label.trim()) {
                    labels.set(value, label.trim());
                  }
                });

                if (labels.size > 0) {
                  optionLabelsByField.set(fieldKey, labels);
                }
              } catch {
                // Ignore single attribute parse errors
              }
            });
          } catch {
            // Ignore metadata parse errors
          }

          const getObjectValue = (obj: any, keys: string[]): any => {
            for (const key of keys) {
              if (obj && obj[key] !== undefined && obj[key] !== null) {
                return obj[key];
              }
            }
            return undefined;
          };

          const isGuid = (value: string): boolean =>
            /^[0-9a-f]{8}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{12}$/i.test(value);

          const normalizeGuid = (value: string): string => {
            const compact = value.replace(/[{}-]/g, '').toLowerCase();
            if (compact.length !== 32) return value.replace(/[{}]/g, '');
            return `${compact.slice(0, 8)}-${compact.slice(8, 12)}-${compact.slice(12, 16)}-${compact.slice(16, 20)}-${compact.slice(20)}`;
          };

          const escapeODataLiteral = (value: string): string => value.replace(/'/g, "''");

          const isLookupField = (fieldName: string): boolean => {
            const fieldType = attributeTypeByField.get((fieldName || '').toLowerCase());
            return fieldType === 'lookup' || fieldType === 'customer' || fieldType === 'owner';
          };

          const parseLookupValue = (raw: string): { entityHint?: string; id: string } | null => {
            const trimmed = String(raw || '').trim();
            if (!trimmed) return null;

            const commaIndex = trimmed.indexOf(',');
            if (commaIndex > 0) {
              const left = trimmed.slice(0, commaIndex).trim();
              const right = trimmed.slice(commaIndex + 1).trim();
              if (isGuid(right)) {
                return { entityHint: left || undefined, id: normalizeGuid(right) };
              }
            }

            const guidMatch = trimmed.match(/[0-9a-fA-F]{8}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{4}-?[0-9a-fA-F]{12}/);
            if (!guidMatch) return null;

            return { id: normalizeGuid(guidMatch[0]) };
          };

          type EntityMetadataInfo = {
            entitySetName: string;
            primaryIdAttribute: string;
            primaryNameAttribute: string | null;
          };

          const entityMetadataCache = new Map<string, EntityMetadataInfo | null>();
          const lookupTargetsByRelationshipCache = new Map<string, string[]>();
          const lookupNameCache = new Map<string, string | null>();

          const getEntityMetadata = async (logicalName: string): Promise<EntityMetadataInfo | null> => {
            const cacheKey = String(logicalName || '').toLowerCase();
            if (!cacheKey) return null;
            if (entityMetadataCache.has(cacheKey)) {
              return entityMetadataCache.get(cacheKey) || null;
            }

            try {
              const metadataUrl =
                `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${escapeODataLiteral(logicalName)}')` +
                `?$select=EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute`;

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
                entityMetadataCache.set(cacheKey, null);
                return null;
              }

              const json = await response.json();
              const entitySetName = typeof json?.EntitySetName === 'string' ? json.EntitySetName : '';
              const primaryIdAttribute = typeof json?.PrimaryIdAttribute === 'string' ? json.PrimaryIdAttribute : '';
              const primaryNameAttribute =
                typeof json?.PrimaryNameAttribute === 'string' ? json.PrimaryNameAttribute : null;

              if (!entitySetName || !primaryIdAttribute) {
                entityMetadataCache.set(cacheKey, null);
                return null;
              }

              const metadata: EntityMetadataInfo = {
                entitySetName,
                primaryIdAttribute,
                primaryNameAttribute,
              };
              entityMetadataCache.set(cacheKey, metadata);
              return metadata;
            } catch {
              entityMetadataCache.set(cacheKey, null);
              return null;
            }
          };

          const getLookupTargetEntities = async (fieldName: string): Promise<string[]> => {
            const fieldKey = (fieldName || '').toLowerCase();
            if (!fieldKey) return [];

            const directTargets = lookupTargetsByField.get(fieldKey);
            if (directTargets && directTargets.length > 0) {
              return directTargets;
            }

            if (lookupTargetsByRelationshipCache.has(fieldKey)) {
              return lookupTargetsByRelationshipCache.get(fieldKey) || [];
            }

            try {
              const relationshipsUrl =
                `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${escapeODataLiteral(entityName)}')` +
                `/ManyToOneRelationships?$select=ReferencedEntity` +
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
                lookupTargetsByRelationshipCache.set(fieldKey, []);
                return [];
              }

              const json = await response.json();
              const targets: string[] = Array.isArray(json?.value)
                ? (json.value as any[])
                    .map((row: any) =>
                      typeof row?.ReferencedEntity === 'string' ? row.ReferencedEntity.trim() : ''
                    )
                    .filter((target: string) => target.length > 0)
                : [];

              const uniqueTargets: string[] = Array.from(new Set<string>(targets));
              lookupTargetsByRelationshipCache.set(fieldKey, uniqueTargets);
              return uniqueTargets;
            } catch {
              lookupTargetsByRelationshipCache.set(fieldKey, []);
              return [];
            }
          };

          const resolveLookupRecordName = async (
            entityLogicalName: string,
            recordId: string
          ): Promise<string | null> => {
            const normalizedId = normalizeGuid(recordId);
            const cacheKey = `${entityLogicalName.toLowerCase()}|${normalizedId.toLowerCase()}`;
            if (lookupNameCache.has(cacheKey)) {
              return lookupNameCache.get(cacheKey) || null;
            }

            const metadata = await getEntityMetadata(entityLogicalName);
            if (!metadata || !metadata.primaryNameAttribute) {
              lookupNameCache.set(cacheKey, null);
              return null;
            }

            try {
              const url =
                `${clientUrl}/api/data/v9.2/${metadata.entitySetName}(${normalizedId})` +
                `?$select=${metadata.primaryNameAttribute}`;

              const response = await fetch(url, {
                method: 'GET',
                headers: {
                  Accept: 'application/json',
                  'OData-MaxVersion': '4.0',
                  'OData-Version': '4.0',
                },
                credentials: 'include',
              });

              if (!response.ok) {
                lookupNameCache.set(cacheKey, null);
                return null;
              }

              const json = await response.json();
              const nameValue = json?.[metadata.primaryNameAttribute];
              const displayName =
                typeof nameValue === 'string' && nameValue.trim().length > 0 ? nameValue.trim() : null;

              lookupNameCache.set(cacheKey, displayName);
              return displayName;
            } catch {
              lookupNameCache.set(cacheKey, null);
              return null;
            }
          };

          const resolveLookupDisplayValue = async (
            fieldName: string,
            rawValue: string
          ): Promise<string | null> => {
            const parsed = parseLookupValue(rawValue);
            if (!parsed) return null;

            const candidateEntities: string[] = [];
            if (parsed.entityHint && parsed.entityHint.trim()) {
              candidateEntities.push(parsed.entityHint.trim());
            }

            const targetsFromMetadata = await getLookupTargetEntities(fieldName);
            candidateEntities.push(...targetsFromMetadata);

            const uniqueCandidates = Array.from(
              new Set(candidateEntities.map((value) => value.toLowerCase()))
            );

            for (const candidate of uniqueCandidates) {
              const displayName = await resolveLookupRecordName(candidate, parsed.id);
              if (displayName) {
                return `${displayName} (${parsed.id})`;
              }
            }

            if (parsed.entityHint) {
              return `${parsed.entityHint} (${parsed.id})`;
            }

            return parsed.id;
          };

          const tryOptionDisplay = (fieldName: string, rawValue: string): string | null => {
            const optionMap = optionLabelsByField.get((fieldName || '').toLowerCase());
            if (!optionMap || !rawValue) return null;

            const normalized = rawValue.trim().toLowerCase();
            const normalizedValue =
              normalized === 'true' ? '1' :
              normalized === 'false' ? '0' :
              rawValue.trim();

            const parts = normalizedValue.split(',').map((part) => part.trim()).filter(Boolean);
            if (parts.length > 1) {
              const displayParts = parts.map((part) => {
                const numeric = Number(part);
                if (!Number.isFinite(numeric)) return part;
                const label = optionMap.get(numeric);
                return label ? `${label} (${numeric})` : String(numeric);
              });
              return displayParts.join('; ');
            }

            const numeric = Number(normalizedValue);
            if (!Number.isFinite(numeric)) return null;
            const label = optionMap.get(numeric);
            return label ? `${label} (${numeric})` : String(numeric);
          };

          const formatAuditRawValue = (value: any): string => {
            if (value === null || value === undefined) return '';
            if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
              return String(value);
            }
            if (Array.isArray(value)) {
              return value.map((entry) => formatAuditRawValue(entry)).filter(Boolean).join('; ');
            }
            if (typeof value === 'object') {
              const idCandidate = getObjectValue(value, ['Id', 'id', 'EntityId', 'entityid', 'Guid', 'guid']);
              if (idCandidate !== undefined) {
                const id = String(idCandidate);
                return isGuid(id) ? normalizeGuid(id) : id;
              }

              const valueCandidate = getObjectValue(value, ['Value', 'value']);
              if (valueCandidate !== undefined) {
                return String(valueCandidate);
              }

              const nameCandidate = getObjectValue(value, ['Name', 'name', 'Label', 'label', 'DisplayName', 'displayName']);
              if (nameCandidate !== undefined) {
                return String(nameCandidate);
              }

              try {
                return JSON.stringify(value);
              } catch {
                return String(value);
              }
            }
            return String(value);
          };

          const isDateField = (fieldName: string): boolean => {
            const t = attributeTypeByField.get((fieldName || '').toLowerCase()) || '';
            return t === 'datetime' || t === 'date' || t === 'datetimeoffset';
          };

          const formatDateValue = (raw: string): string => {
            // Parse ISO-like strings; leave anything else untouched.
            if (!/^\d{4}-\d{2}-\d{2}/.test(raw)) return raw;
            const d = new Date(raw);
            if (isNaN(d.getTime())) return raw;
            return d.toLocaleString();
          };

          const formatAuditDisplayValue = async (fieldName: string, value: any): Promise<string> => {
            if (value === null || value === undefined) return '';

            if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
              const raw = String(value).trim();
              const optionDisplay = tryOptionDisplay(fieldName, raw);
              if (optionDisplay) return optionDisplay;

              if (isLookupField(fieldName)) {
                const lookupDisplay = await resolveLookupDisplayValue(fieldName, raw);
                if (lookupDisplay) return lookupDisplay;
              }

              if (isDateField(fieldName) && typeof value === 'string') {
                return formatDateValue(raw);
              }

              return raw;
            }

            if (Array.isArray(value)) {
              const formattedValues: string[] = [];
              for (const entry of value) {
                const formatted = await formatAuditDisplayValue(fieldName, entry);
                if (formatted) {
                  formattedValues.push(formatted);
                }
              }
              return formattedValues.join('; ');
            }

            if (typeof value === 'object') {
              const nameCandidate = getObjectValue(value, ['Name', 'name', 'Label', 'label', 'DisplayName', 'displayName']);
              const idCandidate = getObjectValue(value, ['Id', 'id', 'EntityId', 'entityid', 'Guid', 'guid']);
              const valueCandidate = getObjectValue(value, ['Value', 'value']);

              if (nameCandidate !== undefined && idCandidate !== undefined) {
                const id = String(idCandidate);
                const normalizedId = isGuid(id) ? normalizeGuid(id) : id;
                return `${String(nameCandidate)} (${normalizedId})`;
              }

              if (nameCandidate !== undefined && valueCandidate !== undefined) {
                const optionDisplay = tryOptionDisplay(fieldName, String(valueCandidate));
                if (optionDisplay) return optionDisplay;
                return `${String(nameCandidate)} (${String(valueCandidate)})`;
              }

              if (nameCandidate !== undefined) {
                return String(nameCandidate);
              }

              if (idCandidate !== undefined) {
                const id = String(idCandidate);
                const normalized = isGuid(id) ? normalizeGuid(id) : id;
                if (isLookupField(fieldName)) {
                  const lookupDisplay = await resolveLookupDisplayValue(fieldName, normalized);
                  if (lookupDisplay) return lookupDisplay;
                }
                return normalized;
              }

              if (valueCandidate !== undefined) {
                const raw = String(valueCandidate);
                const optionDisplay = tryOptionDisplay(fieldName, raw);
                return optionDisplay || raw;
              }

              try {
                return JSON.stringify(value);
              } catch {
                return String(value);
              }
            }

            return String(value);
          };

          // Fallback action labels for the common codes — overridden by the annotation
          // when present (which covers Audit Settings Changes, Activate, Deactivate, etc.).
          const actionLabelFallback: Record<number, string> = {
            1: 'Create',
            2: 'Update',
            3: 'Delete',
            4: 'Assign',
            5: 'Share',
            6: 'Unshare',
          };

          // Process each audit record
          for (const audit of auditEntities) {
            try {
              const auditId = audit.auditid;

              // Action label: prefer the FormattedValue annotation (covers all action codes
              // declared in metadata, even ones we don't have in the fallback map).
              const actionFormatted =
                audit['action@OData.Community.Display.V1.FormattedValue'];
              const actionName: string =
                (typeof actionFormatted === 'string' && actionFormatted.trim()) ||
                actionLabelFallback[audit.action as number] ||
                `Action ${audit.action}`;

              // User name: prefer the formatted-value annotation; fall back to systemuser fetch.
              const userId: string = audit._userid_value || '';
              const userFormatted =
                audit['_userid_value@OData.Community.Display.V1.FormattedValue'];
              const changedBy = await resolveUserName(
                userId,
                typeof userFormatted === 'string' ? userFormatted : undefined
              );

              const changedOn: string = audit.createdon;

              // Try to parse changedata (newer Dataverse format).
              let renderedFromChangeData = false;
              if (audit.changedata) {
                try {
                  const changeData = JSON.parse(audit.changedata);

                  // Newer schema: changedAttributes[]. Some older audits use `attributes`.
                  const changes: any[] = Array.isArray(changeData?.changedAttributes)
                    ? changeData.changedAttributes
                    : Array.isArray(changeData?.attributes)
                    ? changeData.attributes
                    : [];

                  for (const change of changes) {
                    const logicalName = String(change.logicalName || change.LogicalName || change.name || '');
                    if (!logicalName) continue;
                    const oldRawValue = formatAuditRawValue(change.oldValue ?? change.OldValue);
                    const newRawValue = formatAuditRawValue(change.newValue ?? change.NewValue);
                    const oldDisplayValue = await formatAuditDisplayValue(
                      logicalName,
                      change.oldValue ?? change.OldValue
                    );
                    const newDisplayValue = await formatAuditDisplayValue(
                      logicalName,
                      change.newValue ?? change.NewValue
                    );

                    auditRecords.push({
                      auditId: `${auditId}_${logicalName}`,
                      action: actionName,
                      fieldName: logicalName,
                      oldValue: oldDisplayValue,
                      newValue: newDisplayValue,
                      rollbackOldValue: oldRawValue,
                      rollbackNewValue: newRawValue,
                      changedBy,
                      changedOn,
                    });
                    renderedFromChangeData = true;
                  }
                } catch (parseError: any) {
                  console.warn(
                    `[D365 Helper] Failed to parse audit.changedata for audit ${auditId}:`,
                    parseError?.message || parseError
                  );
                }
              }

              if (renderedFromChangeData) continue;

              // Fallback path: no parseable changedata. Try `attributemask` (CSV of column
              // numbers) so users at least see "an update touched N attributes" rather
              // than nothing useful.
              const attributeMask: string =
                typeof audit.attributemask === 'string' ? audit.attributemask : '';
              if (attributeMask.trim()) {
                const count = attributeMask.split(',').filter((s: string) => s.trim()).length;
                auditRecords.push({
                  auditId,
                  action: actionName,
                  fieldName: count > 0 ? `${count} attribute${count === 1 ? '' : 's'} changed` : 'Record',
                  oldValue: '',
                  newValue: actionName,
                  changedBy,
                  changedOn,
                });
                continue;
              }

              // Last resort: just show the action.
              auditRecords.push({
                auditId,
                action: actionName,
                fieldName: audit.operation ? `Operation: ${audit.operation}` : 'Record',
                oldValue: '',
                newValue: actionName,
                changedBy,
                changedOn,
              });
            } catch (recordError: any) {
              console.warn('D365 Helper: Error processing audit record:', recordError);
            }
          }

          result = {
            records: auditRecords,
            recordName,
            entityName,
            totalAuditEntries: auditEntities.length,
            truncated: auditTruncated,
            maxRecords: MAX_AUDIT_RECORDS,
          };
        } catch (error: any) {
          console.error('D365 Helper: Failed to retrieve audit history', error);

          // Extract meaningful error message
          let errorMessage = 'Failed to retrieve audit history. Ensure audit logging is enabled and you have permission to read audit records.';

          if (error) {
            if (typeof error === 'string') {
              errorMessage = error;
            } else if (error.message) {
              errorMessage = error.message;
            } else if (error.error && error.error.message) {
              errorMessage = error.error.message;
            } else if (error.statusText) {
              errorMessage = `Error: ${error.statusText}`;
            } else {
              try {
                errorMessage = JSON.stringify(error);
              } catch (e) {
                errorMessage = 'An unknown error occurred while retrieving audit history.';
              }
            }
          }

          result = {
            records: [],
            error: errorMessage
          };
        }
        break;

      case 'GET_SYSTEM_USERS':
        try {
          // Query enabled system users (not disabled, not application users)
          const usersQuery = '/api/data/v9.2/systemusers?$select=systemuserid,fullname,domainname,internalemailaddress&$filter=isdisabled eq false and accessmode ne 4&$orderby=fullname asc&$top=500';
          
          const usersResponse = await fetch(usersQuery);
          
          if (!usersResponse.ok) {
            throw new Error(`Failed to fetch users: ${usersResponse.statusText}`);
          }
          
          const usersData = await usersResponse.json();
          
          result = {
            users: usersData.value || []
          };
        } catch (error: any) {
          console.error('[D365 Helper Injected] Error fetching system users:', error);
          result = {
            users: [],
            error: error.message || 'Failed to fetch system users'
          };
        }
        break;

      case 'SET_IMPERSONATION':
        try {
          const { userId, fullname, domainname } = data;
          if (!userId) {
            throw new Error('User ID is required');
          }
          
          // Store impersonation state in window
          (window as any).__d365ImpersonatedUser = {
            systemuserid: userId,
            fullname: fullname,
            domainname: domainname
          };
          
          result = { success: true, user: (window as any).__d365ImpersonatedUser };
        } catch (error: any) {
          console.error('D365 Helper: Error setting impersonation:', error);
          result = { success: false, error: error.message };
        }
        break;

      case 'CLEAR_IMPERSONATION':
        try {
          const previousUser = (window as any).__d365ImpersonatedUser;
          (window as any).__d365ImpersonatedUser = null;
          result = { success: true, previousUser: previousUser };
        } catch (error: any) {
          console.error('D365 Helper: Error clearing impersonation:', error);
          result = { success: false, error: error.message };
        }
        break;

      case 'GET_IMPERSONATION_STATUS':
        result = {
          isImpersonating: !!(window as any).__d365ImpersonatedUser,
          user: (window as any).__d365ImpersonatedUser || null
        };
        break;

      case 'GET_ACTIVE_PROCESSES': {
        try {
          const targetEntity =
            (typeof data?.entityName === 'string' && data.entityName.trim()) ||
            (Xrm?.Page?.data?.entity?.getEntityName?.() ?? '');
          if (!targetEntity) {
            result = { entityName: '', processes: [], error: 'Could not determine entity name.' };
            break;
          }

          const escForOData = (v: string) => v.replace(/'/g, "''");
          const orgUrl = Xrm.Utility.getGlobalContext().getClientUrl();

          const url =
            `${orgUrl}/api/data/v9.2/workflows` +
            `?$select=workflowid,name,category,statecode,statuscode,mode,description,` +
            `triggeroncreate,triggerondelete,createdon,modifiedon,_ownerid_value` +
            `&$filter=primaryentity eq '${escForOData(targetEntity)}' and type eq 1` +
            `&$orderby=name asc&$top=500`;

          const response = await fetch(url, {
            method: 'GET',
            headers: {
              Accept: 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
              Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
            },
            credentials: 'include',
          });

          if (!response.ok) {
            const text = await response.text().catch(() => '');
            throw new Error(`HTTP ${response.status}: ${text.slice(0, 200) || response.statusText}`);
          }

          const json = await response.json();
          const rows: any[] = Array.isArray(json?.value) ? json.value : [];

          const categoryLabels: Record<number, string> = {
            0: 'Workflow (Classic)',
            1: 'Dialog',
            2: 'Business Rule',
            3: 'Action',
            4: 'Business Process Flow',
            5: 'Modern Flow',
          };

          const processes = rows.map((row: any) => {
            const isActivated = row.statecode === 1;
            const triggers: string[] = [];
            if (row.triggeroncreate) triggers.push('Create');
            const triggerOnUpdate = row['triggeronupdateattributelist'];
            if (typeof triggerOnUpdate === 'string' && triggerOnUpdate.trim()) triggers.push('Update');
            if (row.triggerondelete) triggers.push('Delete');

            return {
              id: row.workflowid,
              name: row.name || '(unnamed)',
              category: row.category,
              categoryLabel: categoryLabels[row.category] || `Category ${row.category}`,
              mode: row.mode,
              modeLabel: row.mode === 1 ? 'Real-time' : 'Background',
              statecode: row.statecode,
              statuscode: row.statuscode,
              isActivated,
              triggers,
              ownerName: row['_ownerid_value@OData.Community.Display.V1.FormattedValue'] || '',
              modifiedOn: row.modifiedon || row.createdon,
              description: row.description || '',
            };
          });

          result = { entityName: targetEntity, processes };
        } catch (error: any) {
          console.error('[D365 Helper] GET_ACTIVE_PROCESSES failed', error);
          result = {
            entityName: data?.entityName || '',
            processes: [],
            error: error?.message || 'Failed to load active processes.',
          };
        }
        break;
      }

      case 'TOGGLE_PROCESS': {
        try {
          const id = String(data?.id || '').replace(/[{}]/g, '');
          const activate = !!data?.activate;
          if (!id) throw new Error('Missing process id.');

          const orgUrl = Xrm.Utility.getGlobalContext().getClientUrl();
          const payload = activate
            ? { statecode: 1, statuscode: 2 }
            : { statecode: 0, statuscode: 1 };

          const response = await fetch(`${orgUrl}/api/data/v9.2/workflows(${id})`, {
            method: 'PATCH',
            headers: {
              Accept: 'application/json',
              'Content-Type': 'application/json; charset=utf-8',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
              'If-Match': '*',
            },
            credentials: 'include',
            body: JSON.stringify(payload),
          });

          if (!response.ok) {
            const text = await response.text().catch(() => '');
            let parsed = '';
            try {
              const j = JSON.parse(text);
              parsed = j?.error?.message || '';
            } catch {
              parsed = text.slice(0, 300);
            }
            throw new Error(parsed || `HTTP ${response.status} ${response.statusText}`);
          }

          result = { success: true };
        } catch (error: any) {
          result = { success: false, error: error?.message || 'Toggle failed.' };
        }
        break;
      }

      case 'GET_PLUGIN_STEPS': {
        try {
          const targetEntity =
            (typeof data?.entityName === 'string' && data.entityName.trim()) ||
            (Xrm?.Page?.data?.entity?.getEntityName?.() ?? '');
          if (!targetEntity) {
            result = { entityName: '', steps: [], error: 'Could not determine entity name.' };
            break;
          }

          const escForOData = (v: string) => v.replace(/'/g, "''");
          const orgUrl = Xrm.Utility.getGlobalContext().getClientUrl();

          const url =
            `${orgUrl}/api/data/v9.2/sdkmessageprocessingsteps` +
            `?$select=sdkmessageprocessingstepid,name,description,mode,stage,rank,statecode,statuscode,filteringattributes` +
            `&$expand=sdkmessageid($select=name),plugintypeid($select=name,assemblyname,typename),sdkmessagefilterid($select=primaryobjecttypecode)` +
            `&$filter=sdkmessagefilterid/primaryobjecttypecode eq '${escForOData(targetEntity)}'` +
            `&$orderby=stage asc,rank asc&$top=500`;

          const response = await fetch(url, {
            method: 'GET',
            headers: {
              Accept: 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
            },
            credentials: 'include',
          });

          if (!response.ok) {
            const text = await response.text().catch(() => '');
            throw new Error(`HTTP ${response.status}: ${text.slice(0, 200) || response.statusText}`);
          }

          const json = await response.json();
          const rows: any[] = Array.isArray(json?.value) ? json.value : [];

          const stageLabels: Record<number, string> = {
            10: 'Pre-validation',
            20: 'Pre-operation',
            40: 'Post-operation',
          };

          const steps = rows.map((row: any) => {
            const assemblyName = row.plugintypeid?.assemblyname || '';
            // OOB plugins live in Microsoft.* assemblies; everything else is custom.
            const isCustom = !!assemblyName && !/^Microsoft\./i.test(assemblyName);
            return {
              id: row.sdkmessageprocessingstepid,
              name: row.name || '(unnamed)',
              description: row.description || '',
              message: row.sdkmessageid?.name || '(unknown message)',
              primaryEntity: row.sdkmessagefilterid?.primaryobjecttypecode || targetEntity,
              stage: row.stage,
              stageLabel: stageLabels[row.stage] || `Stage ${row.stage}`,
              mode: row.mode,
              modeLabel: row.mode === 0 ? 'Sync' : 'Async',
              rank: row.rank ?? 1,
              isEnabled: row.statecode === 0,
              filteringAttributes: row.filteringattributes || '',
              pluginTypeName: row.plugintypeid?.typename || row.plugintypeid?.name || '',
              assemblyName,
              isCustom,
              images: [] as any[],
            };
          });

          // Follow-up fetch: pre/post images for each step.
          // Build a single $filter with OR over step IDs (capped at a reasonable size).
          if (steps.length > 0) {
            try {
              const stepIds = steps.map((s) => s.id);
              const chunkSize = 50;
              const imageMap = new Map<string, any[]>();
              steps.forEach((s) => imageMap.set(s.id, []));

              for (let i = 0; i < stepIds.length; i += chunkSize) {
                const chunk = stepIds.slice(i, i + chunkSize);
                const filterClause = chunk
                  .map((id) => `_sdkmessageprocessingstepid_value eq ${id}`)
                  .join(' or ');
                const imageUrl =
                  `${orgUrl}/api/data/v9.2/sdkmessageprocessingstepimages` +
                  `?$select=sdkmessageprocessingstepimageid,name,entityalias,imagetype,attributes,messagepropertyname,_sdkmessageprocessingstepid_value` +
                  `&$filter=${encodeURIComponent(filterClause)}` +
                  `&$top=1000`;
                const imgResp = await fetch(imageUrl, {
                  method: 'GET',
                  headers: {
                    Accept: 'application/json',
                    'OData-MaxVersion': '4.0',
                    'OData-Version': '4.0',
                  },
                  credentials: 'include',
                });
                if (!imgResp.ok) continue;
                const imgJson = await imgResp.json();
                const imgRows: any[] = Array.isArray(imgJson?.value) ? imgJson.value : [];
                imgRows.forEach((img) => {
                  const stepId = String(img._sdkmessageprocessingstepid_value || '').toLowerCase();
                  const list = imageMap.get(stepId) || imageMap.get(img._sdkmessageprocessingstepid_value);
                  if (!list) return;
                  const imageTypeLabel =
                    img.imagetype === 0 ? 'Pre' : img.imagetype === 1 ? 'Post' : img.imagetype === 2 ? 'Both' : `Type ${img.imagetype}`;
                  list.push({
                    id: img.sdkmessageprocessingstepimageid,
                    name: img.name || '',
                    entityAlias: img.entityalias || '',
                    imageType: img.imagetype,
                    imageTypeLabel,
                    attributes: img.attributes || '',
                    messagePropertyName: img.messagepropertyname || '',
                  });
                });
              }

              steps.forEach((s) => {
                const list = imageMap.get(s.id) || imageMap.get(String(s.id).toLowerCase()) || [];
                s.images = list;
              });
            } catch (imgErr) {
              console.warn('[D365 Helper] Failed to load step images (non-fatal)', imgErr);
            }
          }

          result = { entityName: targetEntity, steps };
        } catch (error: any) {
          console.error('[D365 Helper] GET_PLUGIN_STEPS failed', error);
          result = {
            entityName: data?.entityName || '',
            steps: [],
            error: error?.message || 'Failed to load plugin steps.',
          };
        }
        break;
      }

      case 'TOGGLE_PLUGIN_STEP': {
        try {
          const id = String(data?.id || '').replace(/[{}]/g, '');
          const enable = !!data?.enable;
          if (!id) throw new Error('Missing step id.');

          const orgUrl = Xrm.Utility.getGlobalContext().getClientUrl();
          const payload = enable
            ? { statecode: 0, statuscode: 1 }
            : { statecode: 1, statuscode: 2 };

          const response = await fetch(`${orgUrl}/api/data/v9.2/sdkmessageprocessingsteps(${id})`, {
            method: 'PATCH',
            headers: {
              Accept: 'application/json',
              'Content-Type': 'application/json; charset=utf-8',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
              'If-Match': '*',
            },
            credentials: 'include',
            body: JSON.stringify(payload),
          });

          if (!response.ok) {
            const text = await response.text().catch(() => '');
            let parsed = '';
            try {
              const j = JSON.parse(text);
              parsed = j?.error?.message || '';
            } catch {
              parsed = text.slice(0, 300);
            }
            throw new Error(parsed || `HTTP ${response.status} ${response.statusText}`);
          }

          result = { success: true };
        } catch (error: any) {
          result = { success: false, error: error?.message || 'Toggle failed.' };
        }
        break;
      }

      case 'UPDATE_PLUGIN_STEP_IMAGE': {
        try {
          const id = String(data?.id || '').replace(/[{}]/g, '');
          if (!id) throw new Error('Missing image id.');

          const orgUrl = Xrm.Utility.getGlobalContext().getClientUrl();
          const payload: Record<string, any> = {};
          if (typeof data?.name === 'string') payload.name = data.name;
          if (typeof data?.entityAlias === 'string') payload.entityalias = data.entityAlias;
          if (typeof data?.attributes === 'string') payload.attributes = data.attributes;
          if (typeof data?.imageType === 'number') payload.imagetype = data.imageType;

          if (Object.keys(payload).length === 0) {
            result = { success: true };
            break;
          }

          const response = await fetch(
            `${orgUrl}/api/data/v9.2/sdkmessageprocessingstepimages(${id})`,
            {
              method: 'PATCH',
              headers: {
                Accept: 'application/json',
                'Content-Type': 'application/json; charset=utf-8',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'If-Match': '*',
              },
              credentials: 'include',
              body: JSON.stringify(payload),
            }
          );

          if (!response.ok) {
            const text = await response.text().catch(() => '');
            let parsed = '';
            try {
              const j = JSON.parse(text);
              parsed = j?.error?.message || '';
            } catch {
              parsed = text.slice(0, 300);
            }
            throw new Error(parsed || `HTTP ${response.status} ${response.statusText}`);
          }

          result = { success: true };
        } catch (error: any) {
          result = { success: false, error: error?.message || 'Update image failed.' };
        }
        break;
      }

      case 'DELETE_PLUGIN_STEP_IMAGE': {
        try {
          const id = String(data?.id || '').replace(/[{}]/g, '');
          if (!id) throw new Error('Missing image id.');

          const orgUrl = Xrm.Utility.getGlobalContext().getClientUrl();
          const response = await fetch(
            `${orgUrl}/api/data/v9.2/sdkmessageprocessingstepimages(${id})`,
            {
              method: 'DELETE',
              headers: {
                Accept: 'application/json',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
              },
              credentials: 'include',
            }
          );

          if (!response.ok) {
            const text = await response.text().catch(() => '');
            let parsed = '';
            try {
              const j = JSON.parse(text);
              parsed = j?.error?.message || '';
            } catch {
              parsed = text.slice(0, 300);
            }
            throw new Error(parsed || `HTTP ${response.status} ${response.statusText}`);
          }

          result = { success: true };
        } catch (error: any) {
          result = { success: false, error: error?.message || 'Delete image failed.' };
        }
        break;
      }

      case 'CREATE_PLUGIN_STEP_IMAGE': {
        try {
          const stepId = String(data?.stepId || '').replace(/[{}]/g, '');
          if (!stepId) throw new Error('Missing step id.');

          const orgUrl = Xrm.Utility.getGlobalContext().getClientUrl();
          const payload: Record<string, any> = {
            name: typeof data?.name === 'string' && data.name.trim() ? data.name : 'PreImage',
            entityalias:
              typeof data?.entityAlias === 'string' && data.entityAlias.trim()
                ? data.entityAlias
                : 'PreImage',
            imagetype: typeof data?.imageType === 'number' ? data.imageType : 0,
            attributes: typeof data?.attributes === 'string' ? data.attributes : '',
            messagepropertyname:
              typeof data?.messagePropertyName === 'string' && data.messagePropertyName.trim()
                ? data.messagePropertyName
                : 'Target',
            'sdkmessageprocessingstepid@odata.bind': `/sdkmessageprocessingsteps(${stepId})`,
          };

          const response = await fetch(`${orgUrl}/api/data/v9.2/sdkmessageprocessingstepimages`, {
            method: 'POST',
            headers: {
              Accept: 'application/json',
              'Content-Type': 'application/json; charset=utf-8',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
              Prefer: 'return=representation',
            },
            credentials: 'include',
            body: JSON.stringify(payload),
          });

          if (!response.ok) {
            const text = await response.text().catch(() => '');
            let parsed = '';
            try {
              const j = JSON.parse(text);
              parsed = j?.error?.message || '';
            } catch {
              parsed = text.slice(0, 300);
            }
            throw new Error(parsed || `HTTP ${response.status} ${response.statusText}`);
          }

          const created = await response.json().catch(() => ({}));
          result = { success: true, id: created?.sdkmessageprocessingstepimageid };
        } catch (error: any) {
          result = { success: false, error: error?.message || 'Create image failed.' };
        }
        break;
      }

      case 'GET_PRIVILEGE_DEBUG': {
        try {
          const targetEntity = String(data?.entityName || '').trim();
          const targetId = String(data?.recordId || '').replace(/[{}]/g, '').trim();
          if (!targetEntity || !targetId) {
            throw new Error('Entity name and record GUID are required.');
          }

          const orgUrl = Xrm.Utility.getGlobalContext().getClientUrl();
          const escForOData = (v: string) => v.replace(/'/g, "''");
          const apiBase = `${orgUrl}/api/data/v9.2`;

          const fetchJson = async (url: string, includeFormatted = false) => {
            const headers: Record<string, string> = {
              Accept: 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0',
            };
            if (includeFormatted) {
              // Request all useful annotations: formatted display values + lookup logical name
              // (so we can derive owner type without needing the non-existent `owneridtype` column).
              headers['Prefer'] = 'odata.include-annotations="*"';
            }
            const r = await fetch(url, {
              method: 'GET',
              headers,
              credentials: 'include',
            });
            if (!r.ok) {
              const text = await r.text().catch(() => '');
              throw new Error(`HTTP ${r.status}: ${text.slice(0, 200) || r.statusText}`);
            }
            return r.json();
          };

          // Resolve entity metadata for entity-set name + primary id/name attribute
          const meta = await fetchJson(
            `${apiBase}/EntityDefinitions(LogicalName='${escForOData(targetEntity)}')` +
              `?$select=EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,ObjectTypeCode`
          );
          const entitySetName: string = meta?.EntitySetName || '';
          const primaryIdAttribute: string = meta?.PrimaryIdAttribute || '';
          const primaryNameAttribute: string | null = meta?.PrimaryNameAttribute || null;
          if (!entitySetName) {
            throw new Error(`Could not resolve entity set for "${targetEntity}".`);
          }

          // WhoAmI
          const whoAmI = await fetchJson(`${apiBase}/WhoAmI`);
          const userId: string = (whoAmI?.UserId || '').replace(/[{}]/g, '');
          const buId: string = (whoAmI?.BusinessUnitId || '').replace(/[{}]/g, '');

          // Current user info + their teams + roles
          const userPromise = fetchJson(
            `${apiBase}/systemusers(${userId})?$select=fullname,isdisabled,_businessunitid_value` +
              `&$expand=businessunitid($select=name)`
          );
          const userTeamsPromise = fetchJson(
            `${apiBase}/systemusers(${userId})/teammembership_association?$select=name,teamid&$top=200`
          ).catch(() => ({ value: [] }));

          // Get user's roles directly
          const userRolesPromise = fetchJson(
            `${apiBase}/systemusers(${userId})/systemuserroles_association?$select=roleid,name,_businessunitid_value` +
              `&$expand=businessunitid($select=name)&$top=200`
          ).catch(() => ({ value: [] }));

          // Record fetch (best-effort).
          // ownerid is polymorphic (principal), so we can't $expand fullname/businessunitid on it directly,
          // and `owneridtype` doesn't exist as a queryable column on most entities.
          // We rely on annotations: `Microsoft.Dynamics.CRM.lookuplogicalname` tells us systemuser vs team,
          // and `OData.Community.Display.V1.FormattedValue` gives the owner's display name.
          const recordSelect = [primaryIdAttribute, '_ownerid_value', '_owningbusinessunit_value'];
          if (primaryNameAttribute) recordSelect.push(primaryNameAttribute);
          const recordPromise = fetchJson(
            `${apiBase}/${entitySetName}(${targetId})?$select=${recordSelect.join(',')}`,
            true
          ).catch((e: any) => ({ __error: e?.message || 'Record fetch failed' }));

          // Effective access (RetrievePrincipalAccess).
          // Dataverse expects Target/Principal as @odata.id entity references; using a typed
          // entity object causes the function dispatcher to lose the Target parameter.
          const principalAccessPromise = (async () => {
            if (!Xrm?.WebApi?.execute) {
              throw new Error('Xrm.WebApi.execute is not available.');
            }
            const request: any = {
              Target: { '@odata.id': `${entitySetName}(${targetId})` },
              Principal: { '@odata.id': `systemusers(${userId})` },
              getMetadata: () => ({
                boundParameter: null,
                operationType: 1, // 1 = Function
                operationName: 'RetrievePrincipalAccess',
                parameterTypes: {
                  Target: { typeName: 'mscrm.crmbaseentity', structuralProperty: 5 },
                  Principal: { typeName: 'mscrm.principal', structuralProperty: 5 },
                },
              }),
            };
            const response = await Xrm.WebApi.execute(request);
            if (!response.ok) {
              throw new Error(`HTTP ${response.status} ${response.statusText}`);
            }
            return await response.json();
          })().catch((e: any) => ({ __error: e?.message || 'Principal access failed' }));

          const [userRes, userTeamsRes, userRolesRes, recordRes, accessRes] = await Promise.all([
            userPromise,
            userTeamsPromise,
            userRolesPromise,
            recordPromise,
            principalAccessPromise,
          ]);

          const teams = ((userTeamsRes?.value as any[]) || [])
            .filter((t: any) => !!t?.teamid)
            .map((t: any) => ({ id: t.teamid, name: t.name || '(team)' }));

          const userRoles = ((userRolesRes?.value as any[]) || []).map((r: any) => ({
            roleId: String(r.roleid || ''),
            roleName: String(r.name || ''),
            buName: r.businessunitid?.name || '',
          }));

          // Also fetch role privileges for each role (limit to first 50 roles)
          const depthRankToLabel = (rank: number): string => {
            if (rank >= 8) return 'Organization';
            if (rank >= 4) return 'Deep';
            if (rank >= 2) return 'Business Unit';
            if (rank >= 1) return 'User';
            return 'None';
          };

          // accessrightsmask bit positions per Microsoft:
          // 1=Read, 2=Write, 4=Append, 8=AppendTo, 16=Create, 32=Delete, 65536=Share, 524288=Assign
          const rightLabel = (mask: number): string | null => {
            switch (mask) {
              case 1: return 'Read';
              case 2: return 'Write';
              case 4: return 'Append';
              case 8: return 'AppendTo';
              case 16: return 'Create';
              case 32: return 'Delete';
              case 65536: return 'Share';
              case 524288: return 'Assign';
              default: return null;
            }
          };

          const otc = meta?.ObjectTypeCode;
          let rolesWithPrivs: any[] = [];
          if (otc) {
            const rolesToFetch = userRoles.slice(0, 50);
            const privPromises = rolesToFetch.map((r) =>
              fetchJson(
                `${apiBase}/roles(${r.roleId})/roleprivileges_association?` +
                  `$select=privilegeid,name,accessright` +
                  `&$top=2000`
              ).catch(() => ({ value: [] }))
            );
            // Also fetch roleprivileges with depth via /roleprivilegescollection style — but RoleId-bound endpoint above
            // doesn't return depth. We need a parallel call to retrieve depth per role.
            const depthPromises = rolesToFetch.map((r) =>
              fetchJson(
                `${apiBase}/roleprivilegescollection?$filter=_roleid_value eq ${r.roleId}&$top=5000`
              ).catch(() => ({ value: [] }))
            );

            const [privsResults, depthResults] = await Promise.all([
              Promise.all(privPromises),
              Promise.all(depthPromises),
            ]);

            // Get privilege metadata to know which privileges apply to our entity OTC.
            const allPrivIds = new Set<string>();
            privsResults.forEach((res: any) => {
              ((res?.value as any[]) || []).forEach((p) => allPrivIds.add(String(p.privilegeid)));
            });

            // Fetch privileges that apply to this entity by OTC (privilegeobjecttypecodes table).
            // privilegeid here is a lookup; OData exposes it as _privilegeid_value, so don't
            // restrict $select (some Dataverse versions reject explicit selection of the FK alias).
            const otcPrivsRes = await fetchJson(
              `${apiBase}/privilegeobjecttypecodes?$filter=objecttypecode eq ${otc}&$top=2000`
            ).catch(() => ({ value: [] }));
            const entityPrivIds = new Set<string>(
              ((otcPrivsRes?.value as any[]) || [])
                .map((p: any) =>
                  String(p?._privilegeid_value || p?.privilegeid || p?.PrivilegeId || '').toLowerCase()
                )
                .filter(Boolean)
            );

            // Build role -> privilege matrix
            rolesWithPrivs = rolesToFetch.map((r, idx) => {
              const privs = ((privsResults[idx]?.value as any[]) || [])
                .filter((p) => {
                  const pid = String(p?.privilegeid || '').toLowerCase();
                  // If we couldn't resolve which privileges apply to this entity,
                  // fall back to including everything (better to over-show than hide everything).
                  return entityPrivIds.size === 0 ? true : entityPrivIds.has(pid);
                })
                .map((p) => ({
                  privilegeId: String(p.privilegeid),
                  accessRight: rightLabel(Number(p.accessright)) || `mask ${p.accessright}`,
                }));

              const depthRows = ((depthResults[idx]?.value as any[]) || []) as any[];
              const depthByPriv = new Map<string, number>();
              depthRows.forEach((d: any) => {
                const pid = String(d?._privilegeid_value || '').toLowerCase();
                const mask = Number(d?.privilegedepthmask || 0);
                if (pid && mask) depthByPriv.set(pid, mask);
              });

              const merged = privs.map((p) => {
                const mask = depthByPriv.get(p.privilegeId.toLowerCase()) || 0;
                return {
                  accessRight: p.accessRight,
                  depth: depthRankToLabel(mask),
                  depthRank: mask,
                };
              });

              return {
                roleId: r.roleId,
                roleName: r.roleName,
                businessUnitName: r.buName,
                privileges: merged,
              };
            });
          }

          // Build result objects
          const recordHasError = recordRes && (recordRes as any).__error;
          const accessHasError = accessRes && (accessRes as any).__error;

          const recordInfo: any = {
            entityName: targetEntity,
            recordId: targetId,
          };
          if (!recordHasError) {
            const r: any = recordRes;
            if (primaryNameAttribute) recordInfo.recordName = r?.[primaryNameAttribute] || '';
            // Derive owner type from the lookup-logical-name annotation rather than the
            // (non-existent on most entities) `owneridtype` column.
            const ownerLogical = String(
              r?.['_ownerid_value@Microsoft.Dynamics.CRM.lookuplogicalname'] || ''
            ).toLowerCase();
            const ownerType =
              ownerLogical === 'team' ? 'team' : ownerLogical === 'systemuser' ? 'user' : '';
            recordInfo.ownerType = ownerType;
            recordInfo.ownerName =
              r?.['_ownerid_value@OData.Community.Display.V1.FormattedValue'] || '';
            recordInfo.recordBusinessUnit =
              r?.['_owningbusinessunit_value@OData.Community.Display.V1.FormattedValue'] || '';

            // Follow-up fetch for the owner's BU once we know the owner type.
            const ownerId = String(r?._ownerid_value || '').replace(/[{}]/g, '');
            if (ownerId && (ownerType === 'user' || ownerType === 'team')) {
              try {
                const ownerSet = ownerType === 'team' ? 'teams' : 'systemusers';
                const ownerRow = await fetchJson(
                  `${apiBase}/${ownerSet}(${ownerId})?$select=_businessunitid_value`,
                  true
                );
                recordInfo.ownerBusinessUnit =
                  ownerRow?.['_businessunitid_value@OData.Community.Display.V1.FormattedValue'] || '';
              } catch {
                // Non-fatal — leave ownerBusinessUnit blank.
              }
            }
          }

          const userInfo: any = {
            userId,
            fullName: userRes?.fullname || '(unknown)',
            businessUnitId: buId,
            businessUnitName: userRes?.businessunitid?.name || '',
            isDisabled: !!userRes?.isdisabled,
            teams,
          };

          const accessRights: any = accessHasError
            ? null
            : (() => {
                const flagsStr = String(accessRes?.AccessRights || '');
                const flags = flagsStr.split(/[,\s]+/).filter(Boolean);
                const has = (k: string) => flags.includes(k);
                return {
                  ReadAccess: has('ReadAccess'),
                  WriteAccess: has('WriteAccess'),
                  DeleteAccess: has('DeleteAccess'),
                  AppendAccess: has('AppendAccess'),
                  AppendToAccess: has('AppendToAccess'),
                  AssignAccess: has('AssignAccess'),
                  ShareAccess: has('ShareAccess'),
                  CreateAccess: has('CreateAccess'),
                };
              })();

          // Diagnosis: simple natural-language summary
          const diagnosis: string[] = [];
          if (recordHasError) diagnosis.push(`Could not load the record: ${(recordRes as any).__error}`);
          if (accessHasError) diagnosis.push(`Effective access check failed: ${(accessRes as any).__error}`);
          if (accessRights) {
            const granted = Object.entries(accessRights)
              .filter(([, v]) => v)
              .map(([k]) => k.replace('Access', ''));
            if (granted.length === 0) {
              diagnosis.push('You currently have NO access to this record.');
            } else {
              diagnosis.push(`You have: ${granted.join(', ')}.`);
            }
          }
          if (recordInfo.ownerName && userInfo.fullName) {
            const isOwnedByYou =
              (recordInfo.ownerType === 'user' && recordInfo.ownerName === userInfo.fullName);
            if (isOwnedByYou) diagnosis.push('You are the owner of this record.');
          }
          if (recordInfo.recordBusinessUnit && userInfo.businessUnitName) {
            if (recordInfo.recordBusinessUnit === userInfo.businessUnitName) {
              diagnosis.push('Record is in the same Business Unit as you.');
            } else {
              diagnosis.push(
                `Record BU (${recordInfo.recordBusinessUnit}) differs from your BU (${userInfo.businessUnitName}). ` +
                  `Access may require Deep or Organization-level depth on this entity.`
              );
            }
          }

          result = {
            inputEntityName: targetEntity,
            inputRecordId: targetId,
            effectiveAccess: accessRights || undefined,
            record: recordInfo,
            user: userInfo,
            roles: rolesWithPrivs,
            diagnosis,
          };
        } catch (error: any) {
          console.error('[D365 Helper] GET_PRIVILEGE_DEBUG failed', error);
          result = {
            inputEntityName: data?.entityName || '',
            inputRecordId: data?.recordId || '',
            error: error?.message || 'Privilege check failed.',
          };
        }
        break;
      }

      default:
        // IMPORTANT: Don't throw error for unknown actions!
        // This allows newer versions of the script to handle new actions
        // while old cached versions silently ignore them
        console.warn('[D365 Helper Injected] Unknown action (ignoring):', action, '| Script version:', INJECTED_SCRIPT_VERSION);
        return; // Exit without sending a response - let another script handle it
    }

    // Send response back to content script
    window.dispatchEvent(new CustomEvent('D365_HELPER_RESPONSE', {
      detail: { requestId, success: true, result, _scriptVersion: INJECTED_SCRIPT_VERSION }
    }));

  } catch (error: any) {
    console.error('[D365 Helper Injected] Sending ERROR response for:', action, '| Error:', error.message, '| RequestId:', requestId);
    window.dispatchEvent(new CustomEvent('D365_HELPER_RESPONSE', {
      detail: { requestId, success: false, error: error.message, _scriptVersion: INJECTED_SCRIPT_VERSION }
    }));
  }
});

