// src/declarations.d.ts

declare module "*.png" {

    const value: string;

    export default value;

}

declare module "*.jpg" {

    const value: string;

    export default value;

}

declare module "*.jpeg" {

    const value: string;

    export default value;

}

declare module "*.svg" {

    const value: string;

    export default value;

}

declare module "*.json" {
    const value: any;
    export default value;

}

declare module 'pdfjs-dist/legacy/build/pdf';
declare module 'pdfjs-dist/legacy/build/pdf.worker.entry';
