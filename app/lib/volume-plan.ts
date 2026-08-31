/**
 * Pure administrative-volume planning.  A volume is never a new witness
 * exhibit bundle: AH1 remains AH1 and only the physical volume changes.
 */
export type VolumeItem = { id: string; pages: number };
export type PlannedVolume<T extends VolumeItem> = { number: number; items: T[]; pages: number; oversize: boolean };

export function planVolumes<T extends VolumeItem>(items: T[], pageLimit: number): PlannedVolume<T>[] {
  if (!items.length) return [];
  if (!Number.isFinite(pageLimit) || pageLimit <= 0) {
    return [{ number: 1, items: [...items], pages: items.reduce((sum, item) => sum + Math.max(0, item.pages), 0), oversize: false }];
  }
  const volumes: PlannedVolume<T>[] = [];
  let current: T[] = [];
  let pages = 0;
  const flush = () => {
    if (!current.length) return;
    volumes.push({ number: volumes.length + 1, items: current, pages, oversize: pages > pageLimit });
    current = []; pages = 0;
  };
  for (const item of items) {
    const itemPages = Math.max(1, item.pages);
    // Never split an exhibit merely to satisfy an administrative threshold.
    if (current.length && pages + itemPages > pageLimit) flush();
    current.push(item); pages += itemPages;
    if (itemPages > pageLimit) flush();
  }
  flush();
  return volumes;
}

export function volumeReference(bundle: string, volume: number, start: number, end: number, multiVolume = false) {
  const range = start === end ? `${start}` : `${start}-${end}`;
  return multiVolume || volume > 1 ? `${bundle}/Vol. ${volume}/${range}` : `${bundle}/${range}`;
}
