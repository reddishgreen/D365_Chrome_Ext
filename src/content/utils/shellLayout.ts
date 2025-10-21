const ensureOriginalTop = (shellContainer: HTMLElement): string => {
  const dataset = shellContainer.dataset as DOMStringMap & {
    d365HelperOriginalTop?: string;
  };

  if (!dataset.d365HelperOriginalTop) {
    const computedTop = window.getComputedStyle(shellContainer).top;
    dataset.d365HelperOriginalTop = computedTop === 'auto' ? '0px' : computedTop;
  }

  return dataset.d365HelperOriginalTop!;
};

const parsePixelValue = (value: string | undefined): number => {
  if (!value) {
    return 0;
  }

  const parsed = parseFloat(value);
  return Number.isFinite(parsed) ? parsed : 0;
};

export const setShellContainerOffset = (offset: number): void => {
  const shellContainer = document.getElementById('shell-container') as HTMLElement | null;
  if (!shellContainer) {
    return;
  }

  const originalTop = ensureOriginalTop(shellContainer);
  const numericTop = parsePixelValue(originalTop);

  shellContainer.style.marginTop = '';
  shellContainer.style.top = `${numericTop + offset}px`;
};

export const restoreShellContainerLayout = (): void => {
  const shellContainer = document.getElementById('shell-container') as HTMLElement | null;
  if (!shellContainer) {
    return;
  }

  const dataset = shellContainer.dataset as DOMStringMap & {
    d365HelperOriginalTop?: string;
  };

  const originalTop = dataset.d365HelperOriginalTop;
  if (originalTop && originalTop !== '0px') {
    shellContainer.style.top = originalTop;
  } else {
    shellContainer.style.removeProperty('top');
  }

  shellContainer.style.marginTop = '';
  delete dataset.d365HelperOriginalTop;
};
