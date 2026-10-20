import { afterEach, describe, expect, it, vi } from "vitest";

// The test runner enables cloud features for upstream tests; this file checks
// the fork's default, where they are hidden.
describe("local-only defaults", () => {
  afterEach(() => {
    vi.unstubAllEnvs();
    vi.resetModules();
  });

  async function loadWithCloud(value: "true" | "false") {
    vi.stubEnv("VITE_ANLG_ENABLE_CLOUD", value);
    vi.resetModules();
    const imports = await import("./providers");
    const calendar = await import("~/calendar/components/shared");
    return { imports, calendar };
  }

  it("hides cloud-account imports and calendars by default", async () => {
    const { imports, calendar } = await loadWithCloud("false");

    const ids = imports.MEETING_IMPORT_PROVIDERS.map((p) => p.id);
    expect(ids).not.toContain("microsoft-teams");
    expect(ids).not.toContain("zoom");
    expect(
      imports.MEETING_IMPORT_PROVIDERS.some(
        (p) => p.directImport === "nango-oauth",
      ),
    ).toBe(false);
    expect(ids).toContain("plaud");

    const visible = calendar.PROVIDERS.filter(
      calendar.isCalendarProviderVisible,
    ).map((p) => p.id);
    expect(visible).toEqual(["apple"]);
  });

  it("restores cloud-account options when enabled", async () => {
    const { imports, calendar } = await loadWithCloud("true");

    const ids = imports.MEETING_IMPORT_PROVIDERS.map((p) => p.id);
    expect(ids).toContain("microsoft-teams");
    expect(
      calendar.PROVIDERS.filter(calendar.isCalendarProviderVisible).map(
        (p) => p.id,
      ),
    ).toEqual(["apple", "google", "outlook"]);
  });
});
