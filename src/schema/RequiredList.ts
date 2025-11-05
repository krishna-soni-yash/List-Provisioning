import { spfi, SPFx, SPFI } from "@pnp/sp";

export const RequiredLists = {
    ProjectMetricLogs: "ProjectMetricLogs",
    EmailLogs: "EmailLogs",
    ProjectMetrics: "ProjectMetrics"
};

export const AllRequiredLists: string[] = [
    RequiredLists.ProjectMetricLogs,
    RequiredLists.EmailLogs,
    RequiredLists.ProjectMetrics
];

export function createPnpSpfx(context: any): SPFI {
    return spfi().using(SPFx(context));
}

export async function provisionRequiredLists(sp: SPFI): Promise<void> {
    const { provisionProjectMetricLogs } = await import('./lists/ProjectMetricsLogs');
    const { provisionEmailLogs } = await import('./lists/EmailLogs');
    const { provisionProjectMetrics } = await import('./lists/ProjectMetrics');

    await provisionProjectMetricLogs(sp);
    await provisionEmailLogs(sp);
    await provisionProjectMetrics(sp);
}

export default AllRequiredLists;
