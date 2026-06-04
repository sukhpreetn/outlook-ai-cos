export interface AICOSConfig {
    aiModelClassify: string;
    aiModelBrief: string;
    aiModelDraft: string;
    tenantId: string;
    clientId: string;
    clientSecret: string;
    userEmail: string;
    anthropicApiKey: string;
    storageConnectionString: string;
    oneDriveFolderId: string;
    batchSize: number;
    tier1ConfidenceFloor: number;
    tier2ConfidenceFloor: number;
    draftCapPerRun: number;
    briefingHour: number;
    briefingTz: string;
}
export declare function getConfig(): Promise<AICOSConfig>;
export declare function invalidateConfigCache(): void;
//# sourceMappingURL=Config.d.ts.map