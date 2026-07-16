import JSZip from 'jszip';
import type { OoxmlMetadata } from '../domain/types';

function parseXml(source: string): Document {
  const document = new DOMParser().parseFromString(source, 'application/xml');
  if (document.querySelector('parsererror')) throw new Error('Érvénytelen OOXML XML-rész.');
  return document;
}

function relationshipTarget(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/');
  parts.pop();
  for (const segment of target.split('/')) {
    if (segment === '..') parts.pop();
    else if (segment !== '.') parts.push(segment);
  }
  return parts.join('/');
}

async function readThemeColors(zip: JSZip): Promise<string[]> {
  const themeSource = await zip.file('xl/theme/theme1.xml')?.async('string');
  if (!themeSource) return [];
  const document = parseXml(themeSource);
  const scheme = document.getElementsByTagName('a:clrScheme')[0];
  if (!scheme) return [];
  return Array.from(scheme.children).map((entry) => {
    const color = entry.firstElementChild;
    const value = color?.getAttribute('lastClr') ?? color?.getAttribute('val') ?? '';
    return value ? `#${value.slice(-6).toUpperCase()}` : '';
  });
}

async function readStyleIds(zip: JSZip): Promise<Map<string, number>> {
  const result = new Map<string, number>();
  const workbookPath = 'xl/workbook.xml';
  const workbookSource = await zip.file(workbookPath)?.async('string');
  const relationsSource = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!workbookSource || !relationsSource) return result;

  const relations = new Map<string, string>();
  const relationDocument = parseXml(relationsSource);
  for (const relation of Array.from(relationDocument.getElementsByTagName('Relationship'))) {
    const id = relation.getAttribute('Id');
    const target = relation.getAttribute('Target');
    if (id && target) relations.set(id, relationshipTarget(workbookPath, target));
  }

  const workbookDocument = parseXml(workbookSource);
  for (const sheet of Array.from(workbookDocument.getElementsByTagName('sheet'))) {
    const name = sheet.getAttribute('name');
    const relationId = sheet.getAttribute('r:id');
    const target = relationId ? relations.get(relationId) : undefined;
    const sheetSource = target ? await zip.file(target)?.async('string') : undefined;
    if (!name || !sheetSource) continue;
    const sheetDocument = parseXml(sheetSource);
    for (const cell of Array.from(sheetDocument.getElementsByTagName('c'))) {
      const address = cell.getAttribute('r');
      const style = cell.getAttribute('s');
      if (address && style !== null) result.set(`${name}!${address}`, Number(style));
    }
  }
  return result;
}

export async function extractOoxmlMetadata(buffer: ArrayBuffer): Promise<OoxmlMetadata> {
  const zip = await JSZip.loadAsync(buffer);
  const [themeColors, styleIds] = await Promise.all([readThemeColors(zip), readStyleIds(zip)]);
  return { themeColors, styleIds };
}
