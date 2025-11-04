// This script runs in the page context and has access to window.Xrm
// It communicates with the content script via custom events

// Store original visibility states
const originalFieldVisibility = new Map<string, boolean>();
const originalSectionVisibility = new Map<string, boolean>();

// Listen for requests from content script
window.addEventListener('D365_HELPER_REQUEST', async (event: any) => {
  const { action, data, requestId } = event.detail;

  try {
    let result: any = null;

    const Xrm = (window as any).Xrm;

    const requiresFormContext = action !== 'GET_PLUGIN_TRACE_LOGS' && action !== 'GET_ENVIRONMENT_ID';

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
          attributes.forEach((attribute: any) => {
            const controls = attribute.controls.get();
            controls.forEach((control: any) => {
              const controlName = control.getName();
              originalFieldVisibility.set(controlName, control.getVisible());
              control.setVisible(true);
            });
          });
        } else {
          // Restore original visibility state
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
          originalSectionVisibility.clear();
          tabs.forEach((tab: any) => {
            const sections = tab.sections.get();
            sections.forEach((section: any) => {
              const sectionName = section.getName();
              originalSectionVisibility.set(sectionName, section.getVisible());
              section.setVisible(true);
            });
          });
        } else {
          // Restore original visibility state
          tabs.forEach((tab: any) => {
            const sections = tab.sections.get();
            sections.forEach((section: any) => {
              const sectionName = section.getName();
              const originalState = originalSectionVisibility.get(sectionName);
              if (originalState !== undefined) {
                section.setVisible(originalState);
              }
            });
          });
          // Clear the saved states after restoration
          originalSectionVisibility.clear();
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
            const currentValue = attr.getValue();

            // Only fill if empty
            if (currentValue === null || currentValue === undefined || currentValue === '') {
              switch (attrType) {
                case 'string':
                case 'memo':
                  attr.setValue('Sample Text');
                  filledCount++;
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
        allAttrs.forEach((attr: any) => {
          const schemaName = attr.getName();
          const controls = attr.controls.get();
          controls.forEach((control: any) => {
            try {
              const controlName = control.getName();

              // Try multiple ways to find the element
              let element = document.getElementById(controlName);

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
                  console.debug('D365 Helper: Error getting control element:', e);
                }
              }

              // Only include controls that are visible AND have visible DOM elements
              const isVisible = control.getVisible();
              if (element && isVisible) {
                controlInfo.push({
                  schemaName: schemaName,
                  controlName: controlName,
                  label: control.getLabel ? control.getLabel() : schemaName,
                  visible: isVisible,
                  elementFound: true
                });
              }
            } catch (e) {
              console.debug('D365 Helper: Error processing control:', e);
            }
          });
        });
        console.log('D365 Helper: Found', controlInfo.length, 'controls');
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
            console.warn('D365 Helper: Failed to process option set attribute', e);
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

          console.warn('D365 Helper: Xrm.Page.data.entity not available');
          console.log('D365 Helper: Current URL:', url);
          console.log('D365 Helper: Is list view:', isListView);

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

        // Debug: Log the structure to understand what's available
        console.log('D365 Helper: Analyzing form for event handlers...');
        console.log('D365 Helper: Xrm.Page.data.entity keys:', Object.keys(Xrm.Page.data.entity));

        // Method 0: Try official Xrm API to get registered event handlers
        try {
          // Check if form-level event methods exist
          const entity = Xrm.Page.data.entity as any;

          // Try to get OnLoad handlers - there's no official "get" method but we can try reflection
          console.log('D365 Helper: Checking for form event registration methods...');

          // OnChange for attributes using official API
          const attributes = Xrm.Page.data.entity.attributes.get();
          attributes.forEach((attr: any) => {
            try {
              // Some D365 versions have addOnChange with registered handlers
              if (typeof attr.addOnChange === 'function') {
                // The function itself might have a reference to registered handlers
                const onChangeFn = attr.addOnChange;
                if ((onChangeFn as any)._handlers) {
                  console.log('D365 Helper: Found _handlers on addOnChange for', attr.getName(), ':', (onChangeFn as any)._handlers);
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
              console.log('D365 Helper: Found _handlers on addOnLoad:', (addOnLoad as any)._handlers);
            }
          }

          // Try accessing OnSave
          if (typeof entity.addOnSave === 'function') {
            const addOnSave = entity.addOnSave;
            if ((addOnSave as any)._handlers) {
              console.log('D365 Helper: Found _handlers on addOnSave:', (addOnSave as any)._handlers);
            }
          }
        } catch (e) {
          console.warn('D365 Helper: Method 0 (Xrm API reflection) failed:', e);
        }

        // Try multiple approaches to get event handlers

        // Method 1: Check _clientApiExecutor._store (new D365 approach)
        try {
          const entity = Xrm.Page.data.entity as any;
          const executor = entity._clientApiExecutor;

          if (executor) {
            console.log('D365 Helper: _clientApiExecutor found, keys:', Object.keys(executor));

            // Check the _store property - this is a Redux store
            if (executor._store) {
              console.log('D365 Helper: executor._store keys:', Object.keys(executor._store));

              const store = executor._store;

              // Get state from Redux store
              if (typeof store.getState === 'function') {
                try {
                  const state = store.getState();
                  console.log('D365 Helper: Redux state keys:', Object.keys(state));

                  // Look for form libraries and event handlers in state
                  if (state.formLibraries) {
                    console.log('D365 Helper: Found formLibraries in state:', state.formLibraries);
                  }

                  if (state.libraries) {
                    console.log('D365 Helper: Found libraries in state:', state.libraries);
                  }

                  if (state.eventHandlers) {
                    console.log('D365 Helper: Found eventHandlers in state:', state.eventHandlers);
                  }

                  if (state.events) {
                    console.log('D365 Helper: Found events in state:', state.events);
                  }

                  if (state.handlers) {
                    console.log('D365 Helper: Found handlers in state:', state.handlers);
                  }

                  // Check for form metadata
                  if (state.form) {
                    console.log('D365 Helper: Found form in state, keys:', Object.keys(state.form));
                    if (state.form.libraries) {
                      console.log('D365 Helper: form.libraries:', state.form.libraries);
                    }
                  }

                  if (state.formData) {
                    console.log('D365 Helper: Found formData in state, keys:', Object.keys(state.formData));
                  }

                  // Check pages state (likely contains form metadata)
                  if (state.pages) {
                    console.log('D365 Helper: Found pages in state, keys:', Object.keys(state.pages));

                    const pages = state.pages;
                    // Pages might be keyed by page ID
                    const pageKeys = Object.keys(pages);
                    if (pageKeys.length > 0) {
                      console.log('D365 Helper: First page key:', pageKeys[0]);
                      const firstPage = pages[pageKeys[0]];
                      if (firstPage) {
                        console.log('D365 Helper: First page keys:', Object.keys(firstPage));

                        // Look for form libraries and event handlers
                        if (firstPage.formLibraries) {
                          console.log('D365 Helper: Found formLibraries in page:', firstPage.formLibraries);
                        }
                        if (firstPage.libraries) {
                          console.log('D365 Helper: Found libraries in page:', firstPage.libraries);
                        }
                        if (firstPage.eventHandlers) {
                          console.log('D365 Helper: Found eventHandlers in page:', firstPage.eventHandlers);
                        }
                        if (firstPage.events) {
                          console.log('D365 Helper: Found events in page:', firstPage.events);
                        }
                        if (firstPage.handlers) {
                          console.log('D365 Helper: Found handlers in page:', firstPage.handlers);
                        }
                        if (firstPage.controls) {
                          console.log('D365 Helper: Found controls in page');
                        }
                        if (firstPage.data) {
                          console.log('D365 Helper: Found data in page, keys:', Object.keys(firstPage.data));
                        }

                        // Check forms in page
                        if (firstPage.forms) {
                          console.log('D365 Helper: Found forms in page, keys:', Object.keys(firstPage.forms));
                          const formKeys = Object.keys(firstPage.forms);
                          if (formKeys.length > 0) {
                            const firstForm = firstPage.forms[formKeys[0]];
                            console.log('D365 Helper: First form keys:', Object.keys(firstForm));

                            if (firstForm.formLibraries) {
                              console.log('D365 Helper: Found formLibraries in form:', firstForm.formLibraries);
                            }
                            if (firstForm.events) {
                              console.log('D365 Helper: Found events in form:', firstForm.events);
                            }
                          }
                        }

                        // Check metadata in page
                        if (firstPage.metadata) {
                          console.log('D365 Helper: Found metadata in page, keys:', Object.keys(firstPage.metadata));
                        }
                      }
                    }
                  }

                  // Check metadata state (might have form definitions)
                  if (state.metadata) {
                    console.log('D365 Helper: Found metadata in state, keys:', Object.keys(state.metadata));

                    const metadata = state.metadata;

                    // Check forms metadata
                    if (metadata.forms) {
                      console.log('D365 Helper: Found forms in metadata, keys:', Object.keys(metadata.forms));

                      // Get current form ID
                      const currentFormId = (Xrm.Page.ui.formSelector?.getCurrentItem()?.getId() || '').replace(/[{}]/g, '');
                      console.log('D365 Helper: Current form ID:', currentFormId);

                      if (currentFormId && metadata.forms[currentFormId]) {
                        const formMetadata = metadata.forms[currentFormId];
                        console.log('D365 Helper: Current form metadata keys:', Object.keys(formMetadata));
                        console.log('D365 Helper: FormMetadata.EventHandlers exists?', !!formMetadata.EventHandlers);
                        console.log('D365 Helper: FormMetadata.eventHandlers exists?', !!formMetadata.eventHandlers);

                        // Extract form libraries (check both PascalCase and camelCase)
                        const formLibraries = formMetadata.FormLibraries || formMetadata.formLibraries;
                        if (formLibraries) {
                          console.log('D365 Helper: Form libraries found:', formLibraries);

                          if (Array.isArray(formLibraries)) {
                            formLibraries.forEach((lib: any) => {
                              const libName = lib.Name || lib.name || lib.LibraryName || lib.libraryName || lib;
                              if (libName && typeof libName === 'string' && !libraries.find(l => l.name === libName)) {
                                libraries.push({ name: libName, order: lib.Order || lib.order || 0 });
                                console.log('D365 Helper: Added library from metadata:', libName);
                              }
                            });
                          }
                        }

                        // Extract event handlers - EventHandlers is an ARRAY of handler objects
                        const eventHandlers = formMetadata.EventHandlers || formMetadata.eventHandlers || formMetadata.events;
                        if (eventHandlers && Array.isArray(eventHandlers)) {
                          console.log('D365 Helper: Event handlers array found with', eventHandlers.length, 'handlers');

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
                              console.log('D365 Helper: Added OnLoad handler:', libraryName, functionName);
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
                              console.log('D365 Helper: Added OnChange handler:', libraryName, functionName, 'for', attributeName);
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
                              console.log('D365 Helper: Added OnSave handler:', libraryName, functionName);
                            }
                          });
                        }

                        // Check for controls with onChange events (use Controls with PascalCase)
                        const controls = formMetadata.Controls || formMetadata.controls;
                        if (controls) {
                          console.log('D365 Helper: Form controls found');

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
                                      console.log('D365 Helper: Added OnChange handler:', libraryName, functionName, 'for', fieldName);
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
                  console.log('D365 Helper: All Redux state top-level keys:', Object.keys(state).join(', '));
                } catch (stateError) {
                  console.warn('D365 Helper: Error accessing Redux state:', stateError);
                }
              }

              // Look for event handlers in store
              if (store._eventHandlers) {
                console.log('D365 Helper: store._eventHandlers:', store._eventHandlers);
              }

              if (store.eventHandlers) {
                console.log('D365 Helper: store.eventHandlers:', store.eventHandlers);
              }

              // Check for onLoad handlers
              if (store.onload || store.onLoad || store.OnLoad) {
                const handlers = store.onload || store.onLoad || store.OnLoad;
                console.log('D365 Helper: Found onload in store:', handlers);
                if (Array.isArray(handlers)) {
                  handlers.forEach((handler: any) => {
                    console.log('D365 Helper: OnLoad handler from store:', handler);
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
                console.log('D365 Helper: Found onsave in store:', handlers);
                if (Array.isArray(handlers)) {
                  handlers.forEach((handler: any) => {
                    console.log('D365 Helper: OnSave handler from store:', handler);
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
                console.log('D365 Helper: Found libraries in store:', libs);
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
              console.log('D365 Helper: executor._eventHandlers:', executor._eventHandlers);
            }

            // Check for registered events
            if (executor._registeredEvents) {
              console.log('D365 Helper: executor._registeredEvents:', executor._registeredEvents);
            }
          }
        } catch (e) {
          console.warn('D365 Helper: Method 1 (_clientApiExecutor) inspection failed:', e);
        }

        // Method 2: Check _formContext (might contain event info)
        try {
          const data = Xrm.Page.data as any;
          const formContext = data._formContext;

          if (formContext) {
            console.log('D365 Helper: _formContext found, keys:', Object.keys(formContext));

            // Deep inspect formContext
            if (formContext._eventHandlers) {
              console.log('D365 Helper: formContext._eventHandlers:', formContext._eventHandlers);
            }

            if (formContext.data) {
              console.log('D365 Helper: formContext.data keys:', Object.keys(formContext.data));
              if ((formContext.data as any)._eventHandlers) {
                console.log('D365 Helper: formContext.data._eventHandlers:', (formContext.data as any)._eventHandlers);
              }
            }
          }
        } catch (e) {
          console.warn('D365 Helper: Method 2 (_formContext) inspection failed:', e);
        }

        // Method 3: Check legacy _eventHandlers property
        try {
          const entity = Xrm.Page.data.entity as any;
          console.log('D365 Helper: Entity _eventHandlers:', entity._eventHandlers);

          if (entity._eventHandlers) {
            // OnLoad
            if (entity._eventHandlers.onload) {
              entity._eventHandlers.onload.forEach((handler: any) => {
                console.log('D365 Helper: Found OnLoad handler:', handler);
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
                console.log('D365 Helper: Found OnSave handler:', handler);
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
          console.warn('D365 Helper: Method 3 (_eventHandlers) failed:', e);
        }

        // Method 4: Check attributes for onChange handlers via _clientApiExecutorAttribute
        try {
          const formAttributes = Xrm.Page.data.entity.attributes.get();
          console.log('D365 Helper: Found', formAttributes.length, 'attributes');

          // Just check first attribute in detail to avoid too much logging
          if (formAttributes.length > 0) {
            const firstAttr = formAttributes[0] as any;
            const fieldName = firstAttr.getName();
            console.log('D365 Helper: Inspecting first attribute:', fieldName);
            console.log('D365 Helper: Attribute keys:', Object.keys(firstAttr));

            if (firstAttr._clientApiExecutorAttribute) {
              console.log('D365 Helper: _clientApiExecutorAttribute keys:', Object.keys(firstAttr._clientApiExecutorAttribute));

              const attrExecutor = firstAttr._clientApiExecutorAttribute;

              // Check for _store in attribute executor
              if (attrExecutor._store) {
                console.log('D365 Helper: Attribute executor._store keys:', Object.keys(attrExecutor._store));

                const attrStore = attrExecutor._store;

                // Get state from attribute's Redux store
                if (typeof attrStore.getState === 'function') {
                  try {
                    const attrState = attrStore.getState();
                    console.log('D365 Helper: Attribute Redux state keys:', Object.keys(attrState));

                    if (attrState.eventHandlers || attrState.events || attrState.onchange) {
                      console.log('D365 Helper: Attribute state:', attrState);
                    }
                  } catch (e) {
                    console.warn('D365 Helper: Error getting attribute state:', e);
                  }
                }

                if (attrStore.onchange || attrStore.onChange || attrStore.OnChange) {
                  console.log('D365 Helper: Attribute has onChange in store!');
                }
              }

              if (attrExecutor._eventHandlers) {
                console.log('D365 Helper: Attribute executor has _eventHandlers:', attrExecutor._eventHandlers);
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
                  console.log('D365 Helper: Found OnChange handler for', fieldName, ':', handler);
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
                    console.log('D365 Helper: Found OnChange handler (via store) for', fieldName, ':', handler);
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
                  console.log('D365 Helper: Found OnChange handler (via executor._eventHandlers) for', fieldName, ':', handler);
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
              console.debug('D365 Helper: Error checking attribute:', e);
            }
          });
        } catch (e) {
          console.warn('D365 Helper: Method 4 (attributes) failed:', e);
        }

        // Method 5: Access form XML metadata (where event handlers are defined)
        try {
          // Try to access form XML from various places
          const data = Xrm.Page.data as any;
          const entity = Xrm.Page.data.entity as any;

          // Check _xrmForm which might contain form XML
          if (entity._xrmForm) {
            console.log('D365 Helper: _xrmForm found, keys:', Object.keys(entity._xrmForm));

            const xrmForm = entity._xrmForm;

            // Look for form XML or form descriptor
            if (xrmForm.FormXml || xrmForm.formXml) {
              const formXml = xrmForm.FormXml || xrmForm.formXml;
              console.log('D365 Helper: Form XML found, length:', formXml?.length);

              if (formXml && typeof formXml === 'string') {
                // Parse the XML to extract libraries and event handlers
                try {
                  const parser = new DOMParser();
                  const xmlDoc = parser.parseFromString(formXml, 'text/xml');

                  // Extract form libraries
                  const formLibraries = xmlDoc.querySelectorAll('Library');
                  console.log('D365 Helper: Found', formLibraries.length, 'libraries in form XML');

                  formLibraries.forEach((lib: any) => {
                    const libName = lib.getAttribute('name');
                    const libOrder = parseInt(lib.getAttribute('order') || '999');
                    if (libName) {
                      console.log('D365 Helper: Library from XML:', libName);
                      if (!libraries.find(l => l.name === libName)) {
                        libraries.push({ name: libName, order: libOrder });
                      }
                    }
                  });

                  // Extract OnLoad event handlers
                  const onLoadEvents = xmlDoc.querySelectorAll('event[name="onload"] Handler, event[name="OnLoad"] Handler');
                  console.log('D365 Helper: Found', onLoadEvents.length, 'OnLoad handlers in XML');

                  onLoadEvents.forEach((handler: any) => {
                    const functionName = handler.getAttribute('functionName');
                    const libraryName = handler.getAttribute('libraryName');
                    const enabled = handler.getAttribute('enabled') !== 'false';

                    if (functionName) {
                      console.log('D365 Helper: OnLoad handler from XML:', libraryName, functionName);
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
                  console.log('D365 Helper: Found', onSaveEvents.length, 'OnSave handlers in XML');

                  onSaveEvents.forEach((handler: any) => {
                    const functionName = handler.getAttribute('functionName');
                    const libraryName = handler.getAttribute('libraryName');
                    const enabled = handler.getAttribute('enabled') !== 'false';

                    if (functionName) {
                      console.log('D365 Helper: OnSave handler from XML:', libraryName, functionName);
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
                  console.log('D365 Helper: Found', onChangeEvents.length, 'OnChange handlers in XML');

                  onChangeEvents.forEach((handler: any) => {
                    const functionName = handler.getAttribute('functionName');
                    const libraryName = handler.getAttribute('libraryName');
                    const enabled = handler.getAttribute('enabled') !== 'false';

                    // Get the parent control to find field name
                    const control = handler.closest('control');
                    const fieldName = control?.getAttribute('datafieldname') || control?.getAttribute('id') || 'Unknown';

                    if (functionName) {
                      console.log('D365 Helper: OnChange handler from XML:', libraryName, functionName, 'for', fieldName);
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
                  console.warn('D365 Helper: Error parsing form XML:', xmlError);
                }
              }
            }

            // Check for FormDescriptor
            if (xrmForm.FormDescriptor || xrmForm.formDescriptor) {
              const desc = xrmForm.FormDescriptor || xrmForm.formDescriptor;
              console.log('D365 Helper: FormDescriptor found:', desc);
            }
          }

          // Try alternate locations for form XML
          if (data._formContext) {
            const formContext = data._formContext;
            console.log('D365 Helper: Checking _formContext for form XML');

            if ((formContext as any).FormXml) {
              console.log('D365 Helper: Found FormXml in _formContext');
            }
          }
        } catch (e) {
          console.warn('D365 Helper: Method 5 (form XML) failed:', e);
        }

        // Method 6: Scan for scripts in DOM and check window object for library namespaces
        try {
          // Look for WebResources (case-insensitive)
          const allScripts = document.querySelectorAll('script[src]');
          console.log('D365 Helper: Found', allScripts.length, 'total script tags');

          const webResourceScripts: any[] = [];
          const customLibNames: string[] = [];

          allScripts.forEach((script: any) => {
            const src = script.getAttribute('src') || '';
            if (src.toLowerCase().includes('webresource')) {
              webResourceScripts.push(script);
              console.log('D365 Helper: WebResource script found:', src);

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

          console.log('D365 Helper: Found', webResourceScripts.length, 'WebResource scripts');
          console.log('D365 Helper: Custom library names:', customLibNames);

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
                console.log('D365 Helper: Adding library from DOM:', libName);
                libraries.push({ name: libName, order: 999 });
              }
            }
          });

          // Try to find library namespaces in window object
          // Common patterns: window.MyNamespace, window.CompanyName, etc.
          const windowKeys = Object.keys(window);
          console.log('D365 Helper: Checking window object for custom namespaces...');

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
            console.log('D365 Helper: Suspected custom namespaces found:', suspectedNamespaces);
            suspectedNamespaces.forEach(ns => {
              const obj = (window as any)[ns];
              const functions = Object.keys(obj).filter(k => typeof obj[k] === 'function');
              console.log(`D365 Helper: Namespace "${ns}" has functions:`, functions.slice(0, 10));
            });
          }
        } catch (e) {
          console.warn('D365 Helper: Method 6 (DOM scan) failed:', e);
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

        console.log('D365 Helper: Final results - Libraries:', libraries.length, 'OnLoad:', onLoadHandlers.length, 'OnChange:', onChangeHandlers.length, 'OnSave:', onSaveHandlers.length);

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
          console.log('D365 Helper: Starting GET_ODATA_FIELDS...');

          // Get entity name and client URL
          const entityLogicalName = Xrm.Page.data.entity.getEntityName();
          const clientUrl = Xrm.Utility.getGlobalContext().getClientUrl();

          console.log('D365 Helper: Entity:', entityLogicalName, 'URL:', clientUrl);

          // Step 1: Fetch basic entity metadata
          const entityMetadataUrl = `${clientUrl}/api/data/v9.2/EntityDefinitions(LogicalName='${entityLogicalName}')?$select=EntitySetName,SchemaName,LogicalName`;

          console.log('D365 Helper: Fetching entity metadata from:', entityMetadataUrl);

          const entityResponse = await fetch(entityMetadataUrl, {
            method: 'GET',
            headers: {
              'Accept': 'application/json',
              'OData-MaxVersion': '4.0',
              'OData-Version': '4.0'
            },
            credentials: 'include'
          });

          console.log('D365 Helper: Entity metadata response status:', entityResponse.status);

          if (!entityResponse.ok) {
            const errorText = await entityResponse.text();
            console.error('D365 Helper: Entity metadata error:', errorText);
            throw new Error(`Failed to fetch entity metadata: ${entityResponse.statusText}`);
          }

          const entityMetadataResult = await entityResponse.json();
          const entitySetName = entityMetadataResult.EntitySetName;
          const entitySchemaName = entityMetadataResult.SchemaName;

          console.log('D365 Helper: Entity metadata success. EntitySetName:', entitySetName);

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

          console.log('D365 Helper: Fetched', attributes.length, 'attributes');

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

            console.log('D365 Helper: Fetched', relationships.length, 'ManyToOne relationships');

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
            console.log('D365 Helper: Could not fetch relationships, continuing without them');
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
            console.log('D365 Helper: Could not fetch lookup targets:', e);
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
              console.log('D365 Helper: Successfully retrieved trace logs using entity name:', entityName);
              break;
            } catch (err: any) {
              console.log('D365 Helper: Failed with entity name:', entityName, err?.message);
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
            messageBlock: log.messageblock || log.messagelog || log.MessageBlock || log.MessageLog
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

      default:
        throw new Error(`Unknown action: ${action}`);
    }

    // Send response back to content script
    window.dispatchEvent(new CustomEvent('D365_HELPER_RESPONSE', {
      detail: { requestId, success: true, result }
    }));

  } catch (error: any) {
    window.dispatchEvent(new CustomEvent('D365_HELPER_RESPONSE', {
      detail: { requestId, success: false, error: error.message }
    }));
  }
});

console.log('D365 Helper injected script loaded');
