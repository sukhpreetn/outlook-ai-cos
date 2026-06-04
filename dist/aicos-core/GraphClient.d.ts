export interface GraphMessage {
    id: string;
    subject: string;
    bodyPreview: string;
    body: {
        content: string;
        contentType: string;
    };
    from: {
        emailAddress: {
            name: string;
            address: string;
        };
    };
    receivedDateTime: string;
    isRead: boolean;
    categories: string[];
    flag: {
        flagStatus: string;
    };
    conversationId: string;
    internetMessageId: string;
    hasAttachments: boolean;
}
export interface GraphEvent {
    id: string;
    subject: string;
    start: {
        dateTime: string;
        timeZone: string;
    };
    end: {
        dateTime: string;
        timeZone: string;
    };
    attendees: Array<{
        emailAddress: {
            name: string;
            address: string;
        };
    }>;
    bodyPreview: string;
    isOnlineMeeting: boolean;
    onlineMeetingUrl: string | null;
}
export declare function fetchNewMessages(_userEmail: string, limit?: number): Promise<GraphMessage[]>;
export declare function autoLabelNewMessages(_userEmail: string): Promise<number>;
export declare function categorizeMessage(_userEmail: string, messageId: string, addCategories: string[], _removeCategories?: string[]): Promise<void>;
export declare function flagMessage(_userEmail: string, messageId: string): Promise<void>;
export declare function createDraftReply(_userEmail: string, messageId: string, body: string): Promise<string>;
export declare function fetchCalendarEvents(_userEmail: string, startDate: Date, endDate: Date): Promise<GraphEvent[]>;
export declare function createTask(_userEmail: string, title: string, body: string, dueDate?: Date, importance?: 'low' | 'normal' | 'high'): Promise<string>;
export declare function appendExcelRow(_userEmail: string, workbookId: string, sheetName: string, values: (string | number | boolean)[]): Promise<void>;
export declare function getExcelRows(_userEmail: string, workbookId: string, sheetName: string): Promise<(string | number | boolean)[][]>;
export declare function createMailSubscription(_userEmail: string, notificationUrl: string): Promise<string>;
//# sourceMappingURL=GraphClient.d.ts.map