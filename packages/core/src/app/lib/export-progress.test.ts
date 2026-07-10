import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { exportSlideAsPdf } from './export-pdf';
import { exportSlideAsImagePptx } from './export-pptx';
import type { SlideModule } from './sdk';

const mocks = vi.hoisted(() => ({
  createRoot: vi.fn(),
  isFrameAnimationSettled: vi.fn(),
  render: vi.fn(),
  strToU8: vi.fn(),
  toBlob: vi.fn(),
  unmount: vi.fn(),
  waitForDataWaitfor: vi.fn(),
  waitForFonts: vi.fn(),
  zipSync: vi.fn(),
}));

vi.mock('react-dom/client', () => ({ createRoot: mocks.createRoot }));
vi.mock('html-to-image', () => ({ toBlob: mocks.toBlob }));
vi.mock('fflate', () => ({ strToU8: mocks.strToU8, zipSync: mocks.zipSync }));
vi.mock('./print-ready', () => ({
  isFrameAnimationSettled: mocks.isFrameAnimationSettled,
  waitForDataWaitfor: mocks.waitForDataWaitfor,
  waitForFonts: mocks.waitForFonts,
}));

class FakeElement {
  readonly attributes = new Map<string, string>();
  readonly children: FakeElement[] = [];
  readonly style = { setProperty: vi.fn() } as unknown as CSSStyleDeclaration;
  className = '';
  download = '';
  href = '';
  id = '';
  parent: FakeElement | null = null;
  rel = '';
  textContent = '';

  constructor(
    readonly tagName: string,
    private readonly onClick?: () => void,
  ) {}

  appendChild<T extends FakeElement>(child: T): T {
    child.parent = this;
    this.children.push(child);
    return child;
  }

  click(): void {
    this.onClick?.();
  }

  querySelectorAll<T extends Element>(): NodeListOf<T> {
    return [] as unknown as NodeListOf<T>;
  }

  remove(): void {
    if (!this.parent) return;
    this.parent.children.splice(this.parent.children.indexOf(this), 1);
    this.parent = null;
  }

  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
  }
}

class FakeDocument {
  readonly body = new FakeElement('body');
  readonly head = new FakeElement('head');
  title = 'Original title';

  constructor(private readonly onDownload: () => void) {}

  createElement(tagName: string): FakeElement {
    return new FakeElement(tagName, tagName === 'a' ? this.onDownload : undefined);
  }
}

function installDom(onDownload = () => {}): {
  document: FakeDocument;
  window: EventTarget & { print: () => void };
} {
  const document = new FakeDocument(onDownload);
  const window = new EventTarget() as EventTarget & { print: () => void };
  window.print = vi.fn();

  vi.stubGlobal('document', document);
  vi.stubGlobal('window', window);
  vi.stubGlobal('requestAnimationFrame', (callback: FrameRequestCallback) => {
    callback(0);
    return 1;
  });
  vi.stubGlobal(
    'getComputedStyle',
    () =>
      ({
        backgroundImage: 'none',
        backgroundPosition: '',
        backgroundRepeat: '',
        backgroundSize: '',
      }) as CSSStyleDeclaration,
  );

  return { document, window };
}

function slide(): SlideModule {
  return { default: [() => null] };
}

beforeEach(() => {
  vi.resetAllMocks();
  mocks.createRoot.mockReturnValue({ render: mocks.render, unmount: mocks.unmount });
  mocks.isFrameAnimationSettled.mockReturnValue(true);
  mocks.strToU8.mockReturnValue(new Uint8Array());
  mocks.toBlob.mockResolvedValue(new Blob(['image']));
  mocks.waitForDataWaitfor.mockResolvedValue(undefined);
  mocks.waitForFonts.mockResolvedValue(undefined);
  mocks.zipSync.mockReturnValue(new Uint8Array([1]));
});

afterEach(() => {
  vi.unstubAllGlobals();
});

describe('export progress completion', () => {
  it('reports PPTX completion once after the download and cleans up', async () => {
    const events: string[] = [];
    const { document } = installDom(() => events.push('download'));

    await exportSlideAsImagePptx(slide(), 'deck', (progress) => events.push(progress.phase));

    expect(events.slice(-2)).toEqual(['download', 'done']);
    expect(events.filter((event) => event === 'done')).toHaveLength(1);
    expect(mocks.unmount).toHaveBeenCalledOnce();
    expect(document.body.children).toHaveLength(0);
    expect(document.head.children).toHaveLength(0);
  });

  it('does not report PPTX completion when capture fails and still cleans up', async () => {
    const error = new Error('capture failed');
    const onProgress = vi.fn();
    const { document } = installDom();
    mocks.toBlob.mockRejectedValue(error);

    await expect(exportSlideAsImagePptx(slide(), 'deck', onProgress)).rejects.toBe(error);

    expect(onProgress).not.toHaveBeenCalledWith(
      expect.objectContaining({ phase: 'done', percent: 100 }),
    );
    expect(mocks.unmount).toHaveBeenCalledOnce();
    expect(document.body.children).toHaveLength(0);
    expect(document.head.children).toHaveLength(0);
  });

  it('reports PDF completion once after printing and cleans up', async () => {
    const events: string[] = [];
    const { document, window } = installDom();
    window.print = vi.fn(() => {
      events.push('print');
      window.dispatchEvent(new Event('afterprint'));
    });

    await exportSlideAsPdf(slide(), 'deck', (progress) => events.push(progress.phase));

    expect(events.slice(-2)).toEqual(['print', 'done']);
    expect(events.filter((event) => event === 'done')).toHaveLength(1);
    expect(mocks.unmount).toHaveBeenCalledOnce();
    expect(document.title).toBe('Original title');
    expect(document.body.children).toHaveLength(0);
    expect(document.head.children).toHaveLength(0);
  });

  it('does not report PDF completion when printing fails and still cleans up', async () => {
    const error = new Error('print failed');
    const onProgress = vi.fn();
    const { document, window } = installDom();
    window.print = vi.fn(() => {
      window.dispatchEvent(new Event('afterprint'));
      throw error;
    });

    await expect(exportSlideAsPdf(slide(), 'deck', onProgress)).rejects.toBe(error);

    expect(onProgress).not.toHaveBeenCalledWith(
      expect.objectContaining({ phase: 'done', percent: 100 }),
    );
    expect(mocks.unmount).toHaveBeenCalledOnce();
    expect(document.title).toBe('Original title');
    expect(document.body.children).toHaveLength(0);
    expect(document.head.children).toHaveLength(0);
  });
});
