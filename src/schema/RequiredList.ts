import { spfi, SPFx, SPFI } from "@pnp/sp";

export const RequiredLists = {
    ProjectMetricLogs: "ProjectMetricLogs",
    EmailLogs: "EmailLogs",
    ProjectMetrics: "ProjectMetrics",
    Test: "Test"
};

export const AllRequiredLists: string[] = [
    RequiredLists.ProjectMetricLogs,
    RequiredLists.EmailLogs,
    RequiredLists.ProjectMetrics,
    RequiredLists.Test
];

export function createPnpSpfx(context: any): SPFI {
    return spfi().using(SPFx(context));
}

export async function provisionRequiredLists(sp: SPFI): Promise<void> {
    //const { provisionProjectMetricLogs } = await import('./lists/ProjectMetricsLogs');
    //const { provisionEmailLogs } = await import('./lists/EmailLogs');
    //const { provisionProjectMetrics } = await import('./lists/ProjectMetrics');
    const { provisionTest } = await import('./lists/Test');

    //await provisionProjectMetricLogs(sp);
    //await provisionEmailLogs(sp);
    //await provisionProjectMetrics(sp);
    await provisionTest(sp);
}

export default AllRequiredLists;
