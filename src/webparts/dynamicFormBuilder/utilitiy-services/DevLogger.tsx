export class DevLogger {
    private static isDevMode: boolean = process.env.NODE_ENV === "development";

    public static info(message: string, data?: any): void {
        if (this.isDevMode) {
            console.log(`%cINFO: ${message}`, "color: blue; font-weight: bold;", data || "");
        }
    }

    public static warn(message: string, data?: any): void {
        if (this.isDevMode) {
            console.warn(`%cWARNING: ${message}`, "color: orange; font-weight: bold;", data || "");
        }
    }

    public static error(message: string, data?: any): void {
        console.error(`%cERROR: ${message}`, "color: red; font-weight: bold;", data || "");
    }

    public static debug(message: string, data?: any): void {
        if (this.isDevMode) {
            console.log(`%cDEBUG: ${message}`, "color: green; font-weight: bold;", data || "");
        }
    }
}
