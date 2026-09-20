import { DataLakeEventBus } from "../src/customEvents";

describe("Data Lake Tables custom events", () => {
    it("emits immutable event payloads and supports unsubscribe", () => {
        const bus = new DataLakeEventBus();
        const events: string[] = [];
        const unsubscribe = bus.on((event) => events.push(`${event.command}:${String(event.payload.name ?? "")}`));
        bus.emit("layoutSaved", { name: "Regional review" });
        unsubscribe();
        bus.emit("layoutRemoved", { name: "Regional review" });
        expect(events).toEqual(["layoutSaved:Regional review"]);
    });

    it("isolates extension listener failures", () => {
        const bus = new DataLakeEventBus();
        const events: string[] = [];
        bus.on(() => { throw new Error("extension failure"); });
        bus.on((event) => events.push(event.command));
        expect(() => bus.emit("filtersChanged")).not.toThrow();
        expect(events).toEqual(["filtersChanged"]);
    });
});
