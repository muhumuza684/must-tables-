export type DataLakeCommand =
    | "layoutApplied"
    | "layoutSaved"
    | "layoutRemoved"
    | "filtersChanged"
    | "exportRequested"
    | "selectionChanged"
    | "copyRequested"
    | "columnGroupToggled";

export interface DataLakeEvent {
    command: DataLakeCommand;
    timestamp: string;
    payload: Record<string, unknown>;
}

export type DataLakeEventListener = (event: DataLakeEvent) => void;

export class DataLakeEventBus {
    private readonly listeners = new Set<DataLakeEventListener>();

    public on(listener: DataLakeEventListener): () => void {
        this.listeners.add(listener);
        return () => this.listeners.delete(listener);
    }

    public emit(command: DataLakeCommand, payload: Record<string, unknown> = {}): void {
        const event: DataLakeEvent = { command, timestamp: new Date().toISOString(), payload: { ...payload } };
        this.listeners.forEach((listener) => {
            try {
                listener(event);
            } catch {
                // Extension code must never be able to break table rendering.
            }
        });
    }

    public clear(): void {
        this.listeners.clear();
    }
}
