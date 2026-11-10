export declare const root: string;
export declare function loadEnvironments(): Promise<Record<string, { origin: string; id: string; displayName: string }>>;
export declare function manifestVersion(packageVersion: string): string;
export declare function renderManifest(envName: string): Promise<string>;
