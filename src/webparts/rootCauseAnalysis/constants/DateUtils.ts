/**
 * Date Utility Functions
 * Contains helper functions for date formatting
 */

/**
 * Pads a number with leading zero if less than 10
 * @param num - The number to pad
 * @returns Padded string representation
 */
const pad = (num: number): string => {
  return num < 10 ? '0' + num : String(num);
};

/**
 * Formats a date string to readable format without T separator
 * @param dateString - The ISO date string to format
 * @returns Formatted date string in MM/DD/YYYY HH:MM:SS format or '-' if invalid
 */
export const formatDate = (dateString: string | undefined): string => {
  if (!dateString) return '-';
  
  try {
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return '-';
    
    // Format as MM/DD/YYYY HH:MM:SS
    const month = pad(date.getMonth() + 1);
    const day = pad(date.getDate());
    const year = date.getFullYear();
    const hours = pad(date.getHours());
    const minutes = pad(date.getMinutes());
    const seconds = pad(date.getSeconds());
    
    return `${month}/${day}/${year} ${hours}:${minutes}:${seconds}`;
  } catch {
    return '-';
  }
};

/**
 * Formats a date string to short date format (date only, no time)
 * @param dateString - The ISO date string to format
 * @returns Formatted date string in MM/DD/YYYY format or '-' if invalid
 */
export const formatDateShort = (dateString: string | undefined): string => {
  if (!dateString) return '-';
  
  try {
    const date = new Date(dateString);
    if (isNaN(date.getTime())) return '-';
    
    // Format as MM/DD/YYYY
    const month = pad(date.getMonth() + 1);
    const day = pad(date.getDate());
    const year = date.getFullYear();
    
    return `${month}/${day}/${year}`;
  } catch {
    return '-';
  }
};
