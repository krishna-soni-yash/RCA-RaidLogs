export const formatResponsibilityValue = (input: any): string => {
  const results: string[] = [];
  const collect = (value: any): void => {
    if (value === null || value === undefined) {
      return;
    }
    if (Array.isArray(value)) {
      value.forEach(collect);
      return;
    }
    if (typeof value === 'object') {
      const email = value?.EMail ?? value?.Email ?? value?.email ?? value?.mail ?? value?.PrimaryEmail;
      if (email) {
        const trimmed = String(email).trim();
        if (trimmed) {
          results.push(trimmed);
        }
        return;
      }
      const title = value?.Title ?? value?.Name ?? value?.text;
      if (title && typeof title === 'string') {
        const sanitized = title.trim();
        if (sanitized) {
          results.push(sanitized);
        }
        return;
      }
    }
    const stringify = String(value).trim();
    if (!stringify) {
      return;
    }
    const pipeSplit = stringify.lastIndexOf('|');
    if (pipeSplit !== -1 && pipeSplit + 1 < stringify.length) {
      const afterPipe = stringify.substring(pipeSplit + 1).trim();
      if (afterPipe) {
        results.push(afterPipe);
        return;
      }
    }
    const hashSplit = stringify.lastIndexOf(';#');
    if (hashSplit !== -1 && hashSplit + 2 <= stringify.length) {
      const afterHash = stringify.substring(hashSplit + 2).trim();
      if (afterHash) {
        results.push(afterHash);
        return;
      }
    }
    results.push(stringify);
  };

  collect(input);
  const filtered = results.filter(Boolean);
  const unique: string[] = [];
  filtered.forEach((value) => {
    if (unique.indexOf(value) === -1) unique.push(value);
  });
  return unique.join(', ');
};

export const formatDateMMDDYYYY = (input: any): string => {
  if (input === null || input === undefined || input === '') return '';
  let dt: Date;
  if (input instanceof Date) dt = input;
  else if (typeof input === 'number') dt = new Date(input);
  else dt = new Date(String(input));
  if (isNaN(dt.getTime())) return '';
  const monthNum = dt.getMonth() + 1;
  const dayNum = dt.getDate();
  const mm = (monthNum < 10 ? '0' : '') + String(monthNum);
  const dd = (dayNum < 10 ? '0' : '') + String(dayNum);
  const yyyy = dt.getFullYear();
  return `${mm}/${dd}/${yyyy}`;
};
