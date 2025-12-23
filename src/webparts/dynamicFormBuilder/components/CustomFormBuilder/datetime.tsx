export const convertUtcToLocalTime = (utcString: string): string => {
    const date = new Date(utcString);
    const options: Intl.DateTimeFormatOptions = {
      hour: '2-digit',
      minute: '2-digit',
      hour12: true,
    };
    return date.toLocaleString(undefined, options);  // `undefined` uses the system's default locale and timezone
  };
export const convertDateToLocal = (dateString: string): string => {
    const [ year, monthStr,day ] = dateString.split("-");
    const date = new Date(Number(year),Number(monthStr)-1, Number(day));
    const options: Intl.DateTimeFormatOptions = {
      day: "numeric",
      month: "short", // Full month name
      year: "numeric",
    };
    return date.toLocaleDateString(undefined, options);
  };
