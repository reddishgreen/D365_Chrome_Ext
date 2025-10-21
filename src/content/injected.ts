// This script runs in the page context and has access to window.Xrm
// It communicates with the content script via custom events

// Listen for requests from content script
window.addEventListener('D365_HELPER_REQUEST', (event: any) => {
  const { action, data, requestId } = event.detail;

  try {
    let result: any = null;

    const Xrm = (window as any).Xrm;

    if (!Xrm || !Xrm.Page) {
      throw new Error('Xrm.Page not available');
    }

    switch (action) {
      case 'GET_RECORD_ID':
        result = Xrm.Page.data.entity.getId().replace(/[{}]/g, '');
        break;

      case 'GET_ENTITY_NAME':
        result = Xrm.Page.data.entity.getEntityName();
        break;

      case 'GET_FORM_ID':
        result = Xrm.Page.ui.formSelector.getCurrentItem().getId();
        break;

      case 'TOGGLE_FIELDS':
        const attributes = Xrm.Page.data.entity.attributes.get();
        attributes.forEach((attribute: any) => {
          const controls = attribute.controls.get();
          controls.forEach((control: any) => {
            control.setVisible(data.show);
          });
        });
        result = { success: true };
        break;

      case 'TOGGLE_SECTIONS':
        const tabs = Xrm.Page.ui.tabs.get();
        tabs.forEach((tab: any) => {
          const sections = tab.sections.get();
          sections.forEach((section: any) => {
            section.setVisible(data.show);
          });
        });
        result = { success: true };
        break;

      case 'GET_SCHEMA_NAMES':
        const attrs = Xrm.Page.data.entity.attributes.get();
        const schemaNames: string[] = [];
        attrs.forEach((attr: any) => {
          schemaNames.push(attr.getName());
        });
        result = schemaNames.sort();
        break;

      case 'UNLOCK_FIELDS':
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

              // Only include controls that have visible DOM elements
              if (element || control.getVisible()) {
                controlInfo.push({
                  schemaName: schemaName,
                  controlName: controlName,
                  label: control.getLabel ? control.getLabel() : schemaName,
                  visible: control.getVisible(),
                  elementFound: !!element
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
