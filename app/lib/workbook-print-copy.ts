export const WORKBOOK_FIDELITY_PREFIX = "Fidelity check failed:";

export const WORKBOOK_TOO_WIDE_LABEL = "This Excel sheet is too wide to print on A4";
export const WORKBOOK_TOO_TALL_LABEL = "This Excel sheet is too tall to print on A4";
export const WORKBOOK_CANNOT_PRINT_LABEL = "This Excel sheet cannot be printed as it stands";
export const WORKSHEET_PRINT_NOTE_LABEL = "Worksheet print note";

export function isWorkbookFidelityWarning(warning: string) {
  return warning.startsWith(WORKBOOK_FIDELITY_PREFIX);
}

export type WorkbookPlanCheckCopy = {
  idSuffix: string;
  blocking: boolean;
  label: string;
  detail: string;
};

type ScaleWarning = {
  warning: string;
  scalePercent: number;
  orientation: string;
  minimumPercent: number;
};

type RangeWarning = {
  warning: string;
  range: string;
  orientation: string;
  minimumPercent: number;
};

type ColumnWarning = {
  warning: string;
  column: string;
  orientation: string;
  minimumPercent: number;
};

function fidelityBody(warning: string) {
  return isWorkbookFidelityWarning(warning)
    ? warning.slice(WORKBOOK_FIDELITY_PREFIX.length).trim()
    : warning.trim();
}

function parseScale(warning: string): ScaleWarning | undefined {
  const match = fidelityBody(warning).match(
    /^unbreakable worksheet content needs (\d+)% scale to fit one (landscape|portrait) A4 printable page; the readable minimum is (\d+)%\.?$/i,
  );
  if (!match) return undefined;
  return {
    warning,
    scalePercent: Number(match[1]),
    orientation: match[2].toLowerCase(),
    minimumPercent: Number(match[3]),
  };
}

function parseMergedWide(warning: string): RangeWarning | undefined {
  const match = fidelityBody(warning).match(
    /^merged cell ([A-Z]+\d+:[A-Z]+\d+) is wider than one (landscape|portrait) A4 printable page at the readable (\d+)% minimum scale\.?$/i,
  );
  if (!match) return undefined;
  return {
    warning,
    range: match[1],
    orientation: match[2].toLowerCase(),
    minimumPercent: Number(match[3]),
  };
}

function parseMergedTall(warning: string): RangeWarning | undefined {
  const match = fidelityBody(warning).match(
    /^merged cell ([A-Z]+\d+:[A-Z]+\d+) is taller than one (landscape|portrait) A4 printable page at the readable (\d+)% minimum scale\.?$/i,
  );
  if (!match) return undefined;
  return {
    warning,
    range: match[1],
    orientation: match[2].toLowerCase(),
    minimumPercent: Number(match[3]),
  };
}

function parseColumnWide(warning: string): ColumnWarning | undefined {
  const match = fidelityBody(warning).match(
    /^column ([A-Z]+) is wider than one (landscape|portrait) A4 printable page at the readable (\d+)% minimum scale\.?$/i,
  );
  if (!match) return undefined;
  return {
    warning,
    column: match[1],
    orientation: match[2].toLowerCase(),
    minimumPercent: Number(match[3]),
  };
}

function parseTitleRows(warning: string) {
  const match = fidelityBody(warning).match(
    /^repeated title rows need more than one (landscape|portrait) A4 printable page at the readable (\d+)% minimum scale\.?$/i,
  );
  if (!match) return undefined;
  return { warning, orientation: match[1].toLowerCase(), minimumPercent: Number(match[2]) };
}

function parsePageCount(warning: string) {
  const match = fidelityBody(warning).match(
    /^worksheet pagination predicts (\d+) pages, above the (\d+)-page native Excel safety limit\.?$/i,
  );
  if (!match) return undefined;
  return { warning, pages: Number(match[1]), limit: Number(match[2]) };
}

function joinRanges(ranges: string[]) {
  if (ranges.length === 1) return ranges[0];
  if (ranges.length === 2) return `${ranges[0]} and ${ranges[1]}`;
  return `${ranges.slice(0, -1).join(", ")} and ${ranges[ranges.length - 1]}`;
}

function otherFidelityDetail(sheetName: string, warning: string) {
  const body = fidelityBody(warning);
  if (/Excel print titles are invalid or belong to another worksheet/i.test(body)) {
    return `${sheetName} has Excel print titles that are invalid or belong to another worksheet.`;
  }
  if (/Excel print-title rows are invalid or exceed the worksheet analysis limit/i.test(body)) {
    return `Print-title rows on ${sheetName} are invalid or exceed the analysis limit.`;
  }
  if (/Excel print-title columns are invalid or exceed the worksheet analysis limit/i.test(body)) {
    return `Print-title columns on ${sheetName} are invalid or exceed the analysis limit.`;
  }
  if (/Excel print titles mix row and column coordinates/i.test(body)) {
    return `Print titles on ${sheetName} mix row and column coordinates.`;
  }
  if (/repeated print-title columns are not supported/i.test(body)) {
    return `Repeated print-title columns on ${sheetName} are not supported. Remove them in Excel, or print this worksheet separately.`;
  }
  if (/repeated print-title rows must be a leading prefix/i.test(body)) {
    return `Repeated print-title rows on ${sheetName} must be a leading prefix of the print range.`;
  }
  if (/worksheet comments are configured to print/i.test(body)) {
    return `${sheetName} is set to print comments, which Exhibit Builder cannot include. Turn comment printing off, or print this worksheet separately.`;
  }
  if (/headers or footers contain printable content/i.test(body)) {
    return `${sheetName} has a header or footer. Remove it, or print this worksheet separately.`;
  }
  const printArea = body.match(/^Excel's print area (.+) cuts through merged cell (.+)\.$/i);
  if (printArea) {
    return `The print area ${printArea[1]} on ${sheetName} cuts through merged cells ${printArea[2]}. Expand the print area so the merge stays whole.`;
  }
  if (/comments, notes, or header\/footer pictures/i.test(body)) {
    return `${sheetName} has comments, notes, or header or footer pictures that Exhibit Builder cannot print. Remove them, or print this worksheet separately.`;
  }
  if (/printed row or column headings/i.test(body)) {
    return `${sheetName} is set to print row or column headings. Turn that off, or print this worksheet separately.`;
  }
  if (/drawings or charts/i.test(body)) {
    return `${sheetName} has drawings or charts whose print position cannot be checked. Remove them, or print this worksheet separately.`;
  }
  return `${sheetName}: ${body}`;
}

function scaleAndFitDetail(sheetName: string, scale: ScaleWarning, merged: RangeWarning[], columns: ColumnWarning[]) {
  const mergedText = merged.length ? `merged cells ${joinRanges(merged.map((item) => item.range))}` : "";
  const columnText = columns.length
    ? `${columns.length === 1 ? "column" : "columns"} ${joinRanges(columns.map((item) => item.column))}`
    : "";
  const cause = mergedText && columnText
    ? ` ${mergedText[0].toUpperCase()}${mergedText.slice(1)} and ${columnText} would need the sheet printed at ${scale.scalePercent}%.`
    : mergedText
      ? ` ${mergedText[0].toUpperCase()}${mergedText.slice(1)} would need the sheet printed at ${scale.scalePercent}%.`
      : columnText
        ? ` ${columnText[0].toUpperCase()}${columnText.slice(1)} would need the sheet printed at ${scale.scalePercent}%.`
        : ` Fitting it would require shrinking the print to ${scale.scalePercent}%.`;
  const fix = merged.length
    ? `In Excel, split ${joinRanges(merged.map((item) => item.range))} or narrow the columns in that heading. If this Excel file should not be in the bundle, leave it out.`
    : columns.length
      ? `In Excel, narrow ${joinRanges(columns.map((item) => `column ${item.column}`))}. If this Excel file should not be in the bundle, leave it out.`
      : "Narrow the columns, split any wide merged heading, or leave this Excel file out of the bundle.";
  return `${sheetName} is too wide for ${scale.orientation} A4.${cause} Exhibit Builder stops shrinking at ${scale.minimumPercent}% so the page stays readable.\n\n${fix}`;
}

export function workbookPlanCheckCopy(sheetName: string, warnings: string[]): WorkbookPlanCheckCopy[] {
  const notes: string[] = [];
  const scale: ScaleWarning[] = [];
  const mergedWide: RangeWarning[] = [];
  const mergedTall: RangeWarning[] = [];
  const columns: ColumnWarning[] = [];
  const titleRows: Array<NonNullable<ReturnType<typeof parseTitleRows>>> = [];
  const pageCounts: Array<NonNullable<ReturnType<typeof parsePageCount>>> = [];
  const otherFidelity: string[] = [];

  for (const warning of warnings) {
    if (!isWorkbookFidelityWarning(warning)) {
      notes.push(warning);
      continue;
    }
    const scaleMatch = parseScale(warning);
    if (scaleMatch) {
      scale.push(scaleMatch);
      continue;
    }
    const wide = parseMergedWide(warning);
    if (wide) {
      mergedWide.push(wide);
      continue;
    }
    const tall = parseMergedTall(warning);
    if (tall) {
      mergedTall.push(tall);
      continue;
    }
    const column = parseColumnWide(warning);
    if (column) {
      columns.push(column);
      continue;
    }
    const titles = parseTitleRows(warning);
    if (titles) {
      titleRows.push(titles);
      continue;
    }
    const pages = parsePageCount(warning);
    if (pages) {
      pageCounts.push(pages);
      continue;
    }
    otherFidelity.push(warning);
  }

  const checks: WorkbookPlanCheckCopy[] = [];
  const fitScale = scale[0];
  if (fitScale && (mergedWide.length || columns.length)) {
    checks.push({
      idSuffix: "print-fit",
      blocking: true,
      label: WORKBOOK_TOO_WIDE_LABEL,
      detail: scaleAndFitDetail(sheetName, fitScale, mergedWide, columns),
    });
  } else {
    if (fitScale) {
      checks.push({
        idSuffix: fitScale.warning,
        blocking: true,
        label: WORKBOOK_TOO_WIDE_LABEL,
        detail: `${sheetName} will not fit on one ${fitScale.orientation} A4 page unless Excel shrinks it to ${fitScale.scalePercent}%. Exhibit Builder does not shrink unbreakable rows, columns or merged cells below ${fitScale.minimumPercent}%, because the printed page would be hard to read.\n\nNarrow the columns, split any wide merged heading, or leave this Excel file out of the bundle.`,
      });
    }
    if (mergedWide.length) {
      checks.push({
        idSuffix: mergedWide[0].warning,
        blocking: true,
        label: WORKBOOK_TOO_WIDE_LABEL,
        detail: `Merged cells ${joinRanges(mergedWide.map((item) => item.range))} on ${sheetName} are wider than one ${mergedWide[0].orientation} A4 page at ${mergedWide[0].minimumPercent}% size. Split that merge or narrow the columns so the heading fits.`,
      });
    }
    if (columns.length) {
      checks.push({
        idSuffix: columns[0].warning,
        blocking: true,
        label: WORKBOOK_TOO_WIDE_LABEL,
        detail: `Column ${joinRanges(columns.map((item) => item.column))} on ${sheetName} ${columns.length === 1 ? "is" : "are"} wider than one ${columns[0].orientation} A4 page at ${columns[0].minimumPercent}% size. Narrow ${columns.length === 1 ? "that column" : "those columns"}, or leave this Excel file out of the bundle.`,
      });
    }
  }

  for (const item of mergedTall) {
    checks.push({
      idSuffix: item.warning,
      blocking: true,
      label: WORKBOOK_TOO_TALL_LABEL,
      detail: `Merged cells ${item.range} on ${sheetName} are taller than one ${item.orientation} A4 page at ${item.minimumPercent}% size. Split that merge or reduce the row height so it fits.`,
    });
  }
  for (const item of titleRows) {
    checks.push({
      idSuffix: item.warning,
      blocking: true,
      label: WORKBOOK_TOO_TALL_LABEL,
      detail: `Repeated title rows on ${sheetName} need more than one ${item.orientation} A4 page at ${item.minimumPercent}% size. Reduce those title rows, or leave this Excel file out of the bundle.`,
    });
  }
  for (const item of pageCounts) {
    checks.push({
      idSuffix: item.warning,
      blocking: true,
      label: WORKBOOK_CANNOT_PRINT_LABEL,
      detail: `${sheetName} would print as ${item.pages} pages. Exhibit Builder stops at ${item.limit} pages for a native Excel print. Split the sheet or reduce the print range.`,
    });
  }
  for (const warning of otherFidelity) {
    checks.push({
      idSuffix: warning,
      blocking: true,
      label: WORKBOOK_CANNOT_PRINT_LABEL,
      detail: otherFidelityDetail(sheetName, warning),
    });
  }
  for (const warning of notes) {
    checks.push({
      idSuffix: warning,
      blocking: false,
      label: WORKSHEET_PRINT_NOTE_LABEL,
      detail: `${sheetName}: ${warning}`,
    });
  }
  return checks;
}
