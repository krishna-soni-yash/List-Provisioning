import { SPFI } from "@pnp/sp";
import { RequiredLists } from '../RequiredList';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";

const LIST_TITLE = RequiredLists.ProjectMetrics;

export async function provisionProjectMetrics(sp: SPFI): Promise<void> {
    async function listExists(title: string): Promise<boolean> {
        try {
            await sp.web.lists.getByTitle(title).select("Id")();
            return true;
        } catch (e) {
            return false;
        }
    }

    async function fieldExists(listTitle: string, fieldName: string): Promise<boolean> {
        try {
            await sp.web.lists.getByTitle(listTitle).fields.getByInternalNameOrTitle(fieldName).select("Id")();
            return true;
        } catch (e) {
            return false;
        }
    }

    let list: any;
    const exists = await listExists(LIST_TITLE);
    if (!exists) {
        const ensureResult = await sp.web.lists.ensure(LIST_TITLE, "Project metrics list", 100);
        list = ensureResult.list;
    } else {
        list = sp.web.lists.getByTitle(LIST_TITLE);
    }

    async function ensureField(xml: string, internalName: string) {
        if (await fieldExists(LIST_TITLE, internalName)) {
            return;
        }
        await list.fields.createFieldAsXml(xml, true);
    }

    const logsListInfo: any = await sp.web.lists.getByTitle(RequiredLists.ProjectMetricLogs).select("Id")();
    if (logsListInfo && logsListInfo.Id) {
        let logsListId = logsListInfo.Id as string;
        if (logsListId.charAt(0) !== '{') {
            logsListId = `{${logsListId}}`;
        }
        const versionLookupXml = `<Field Type='Lookup' Name='VersionId' StaticName='VersionId' DisplayName='VersionId' List='${logsListId}' ShowField='ID' />`;
        await ensureField(versionLookupXml, "VersionId");
    }


    // Choice (Yes/No)
    const yesNoChoices = ["Yes", "No"];
    const yesNoChoicesXml = yesNoChoices.map(c => `<CHOICE>${c}</CHOICE>`).join("");

    const applicabilityXml = `<Field Type='Choice' Name='Applicability' StaticName='Applicability' DisplayName='Applicability' Format='Dropdown'><CHOICES>${yesNoChoicesXml}</CHOICES></Field>`;
    await ensureField(applicabilityXml, "Applicability");

    const associatedPpmXml = `<Field Type='Text' Name='AssociatedPPM' StaticName='AssociatedPPM' DisplayName='AssociatedPPM' MaxLength='255' />`;
    await ensureField(associatedPpmXml, "AssociatedPPM");

    const baselineAndRevisionFrequencyXml = `<Field Type='Text' Name='BaselineAndRevisionFrequency' StaticName='BaselineAndRevisionFrequency' DisplayName='BaselineAndRevisionFrequency' MaxLength='255' />`;
    await ensureField(baselineAndRevisionFrequencyXml, "BaselineAndRevisionFrequency");

    const bgXml = `<Field Type='Text' Name='BG' StaticName='BG' DisplayName='BG' MaxLength='255' />`;
    await ensureField(bgXml, "BG");

    const causalAnalysisTriggerXml = `<Field Type='Note' Name='CausalAnalysisTrigger' StaticName='CausalAnalysisTrigger' DisplayName='CausalAnalysisTrigger' NumLines='6' RichText='FALSE' />`;
    await ensureField(causalAnalysisTriggerXml, "CausalAnalysisTrigger");

    const dataCollectionFrequencyXml = `<Field Type='Text' Name='DataCollectionFrequency' StaticName='DataCollectionFrequency' DisplayName='DataCollectionFrequency' MaxLength='255' />`;
    await ensureField(dataCollectionFrequencyXml, "DataCollectionFrequency");

    const dataAnalysisFrequencyXml = `<Field Type='Text' Name='DataAnalysisFrequency' StaticName='DataAnalysisFrequency' DisplayName='DataAnalysisFrequency' MaxLength='255' />`;
    await ensureField(dataAnalysisFrequencyXml, "DataAnalysisFrequency");

    const dataInputXml = `<Field Type='Note' Name='DataInput' StaticName='DataInput' DisplayName='DataInput' NumLines='6' RichText='FALSE' />`;
    await ensureField(dataInputXml, "DataInput");

    const dataSourceXml = `<Field Type='Text' Name='DataSource' StaticName='DataSource' DisplayName='DataSource' MaxLength='255' />`;
    await ensureField(dataSourceXml, "DataSource");

    const goalXml = `<Field Type='Text' Name='Goal' StaticName='Goal' DisplayName='Goal' MaxLength='255' />`;
    await ensureField(goalXml, "Goal");

    const hasSubProcessXml = `<Field Type='Choice' Name='HasSubProcess' StaticName='HasSubProcess' DisplayName='HasSubProcess' Format='Dropdown'><CHOICES>${yesNoChoicesXml}</CHOICES></Field>`;
    await ensureField(hasSubProcessXml, "HasSubProcess");

    const interpretationGuidelinesXml = `<Field Type='Note' Name='InterpretationGuidelines' StaticName='InterpretationGuidelines' DisplayName='InterpretationGuidelines' NumLines='6' RichText='FALSE' />`;
    await ensureField(interpretationGuidelinesXml, "InterpretationGuidelines");

    const isActiveXml = `<Field Type='Boolean' Name='IsActive' StaticName='IsActive' DisplayName='IsActive' />`;
    await ensureField(isActiveXml, "IsActive");

    const lslXml = `<Field Type='Number' Name='LSL' StaticName='LSL' DisplayName='LSL' />`;
    await ensureField(lslXml, "LSL");

    const metricsXml = `<Field Type='Text' Name='Metrics' StaticName='Metrics' DisplayName='Metrics' MaxLength='255' />`;
    await ensureField(metricsXml, "Metrics");

    const metricsFormulaeXml = `<Field Type='Note' Name='MetricsFormulae' StaticName='MetricsFormulae' DisplayName='MetricsFormulae' NumLines='6' RichText='FALSE' />`;
    await ensureField(metricsFormulaeXml, "MetricsFormulae");

    const orgCausalAnalysisTriggerXml = `<Field Type='Note' Name='OrgCausalAnalysisTrigger' StaticName='OrgCausalAnalysisTrigger' DisplayName='OrgCausalAnalysisTrigger' NumLines='6' RichText='FALSE' />`;
    await ensureField(orgCausalAnalysisTriggerXml, "OrgCausalAnalysisTrigger");

    const orgInterpretationGuidelinesXml = `<Field Type='Note' Name='OrgInterpretationGuidelines' StaticName='OrgInterpretationGuidelines' DisplayName='OrgInterpretationGuidelines' NumLines='6' RichText='FALSE' />`;
    await ensureField(orgInterpretationGuidelinesXml, "OrgInterpretationGuidelines");

    const orgStatisticalXml = `<Field Type='Text' Name='OrgStatistical' StaticName='OrgStatistical' DisplayName='OrgStatistical' MaxLength='255' />`;
    await ensureField(orgStatisticalXml, "OrgStatistical");

    const performanceGoalsXml = `<Field Type='Text' Name='PerformanceGoals' StaticName='PerformanceGoals' DisplayName='PerformanceGoals' MaxLength='255' />`;
    await ensureField(performanceGoalsXml, "PerformanceGoals");

    const pgXml = `<Field Type='Text' Name='PG' StaticName='PG' DisplayName='PG' MaxLength='255' />`;
    await ensureField(pgXml, "PG");

    const priorityXml = `<Field Type='Text' Name='Priority' StaticName='Priority' DisplayName='Priority' MaxLength='255' />`;
    await ensureField(priorityXml, "Priority");

    const probabilityOfSuccessThresholdXml = `<Field Type='Text' Name='ProbabilityOfSuccessThreshold' StaticName='ProbabilityOfSuccessThreshold' DisplayName='ProbabilityOfSuccessThreshold' MaxLength='255' />`;
    await ensureField(probabilityOfSuccessThresholdXml, "ProbabilityOfSuccessThreshold");

    const processXml = `<Field Type='Text' Name='Process' StaticName='Process' DisplayName='Process' MaxLength='255' />`;
    await ensureField(processXml, "Process");

    const projectTypeXml = `<Field Type='Text' Name='ProjectType' StaticName='ProjectType' DisplayName='ProjectType' MaxLength='255' />`;
    await ensureField(projectTypeXml, "ProjectType");

    const quantitativeXml = `<Field Type='Text' Name='Quantitative' StaticName='Quantitative' DisplayName='Quantitative' MaxLength='255' />`;
    await ensureField(quantitativeXml, "Quantitative");

    const statisticalXml = `<Field Type='Text' Name='Statistical' StaticName='Statistical' DisplayName='Statistical' MaxLength='255' />`;
    await ensureField(statisticalXml, "Statistical");

    const subApplicabilityXml = `<Field Type='Text' Name='SubApplicability' StaticName='SubApplicability' DisplayName='SubApplicability' MaxLength='255' />`;
    await ensureField(subApplicabilityXml, "SubApplicability");

    const subBaselineAndRevisionFrequencyXml = `<Field Type='Text' Name='SubBaselineAndRevisionFrequency' StaticName='SubBaselineAndRevisionFrequency' DisplayName='SubBaselineAndRevisionFrequency' MaxLength='255' />`;
    await ensureField(subBaselineAndRevisionFrequencyXml, "SubBaselineAndRevisionFrequency");

    const subDataAnalysisFrequencyXml = `<Field Type='Text' Name='SubDataAnalysisFrequency' StaticName='SubDataAnalysisFrequency' DisplayName='SubDataAnalysisFrequency' MaxLength='255' />`;
    await ensureField(subDataAnalysisFrequencyXml, "SubDataAnalysisFrequency");

    const subDataCollectionFrequencyXml = `<Field Type='Text' Name='SubDataCollectionFrequency' StaticName='SubDataCollectionFrequency' DisplayName='SubDataCollectionFrequency' MaxLength='255' />`;
    await ensureField(subDataCollectionFrequencyXml, "SubDataCollectionFrequency");

    const subDataInputXml = `<Field Type='Note' Name='SubDataInput' StaticName='SubDataInput' DisplayName='SubDataInput' NumLines='6' RichText='FALSE' />`;
    await ensureField(subDataInputXml, "SubDataInput");

    const subDataSourceXml = `<Field Type='Text' Name='SubDataSource' StaticName='SubDataSource' DisplayName='SubDataSource' MaxLength='255' />`;
    await ensureField(subDataSourceXml, "SubDataSource");

    const subGoalXml = `<Field Type='Text' Name='SubGoal' StaticName='SubGoal' DisplayName='SubGoal' MaxLength='255' />`;
    await ensureField(subGoalXml, "SubGoal");

    const subLslXml = `<Field Type='Number' Name='SubLSL' StaticName='SubLSL' DisplayName='SubLSL' />`;
    await ensureField(subLslXml, "SubLSL");

    const subMetricsXml = `<Field Type='Text' Name='SubMetrics' StaticName='SubMetrics' DisplayName='SubMetrics' MaxLength='255' />`;
    await ensureField(subMetricsXml, "SubMetrics");

    const subMetricsFormulaeXml = `<Field Type='Note' Name='SubMetricsFormulae' StaticName='SubMetricsFormulae' DisplayName='SubMetricsFormulae' NumLines='6' RichText='FALSE' />`;
    await ensureField(subMetricsFormulaeXml, "SubMetricsFormulae");

    const subprocessXml = `<Field Type='Text' Name='Subprocess' StaticName='Subprocess' DisplayName='Subprocess' MaxLength='255' />`;
    await ensureField(subprocessXml, "Subprocess");

    const subUnitOfMeasureXml = `<Field Type='Text' Name='SubUnitOfMeasure' StaticName='SubUnitOfMeasure' DisplayName='SubUnitOfMeasure' MaxLength='255' />`;
    await ensureField(subUnitOfMeasureXml, "SubUnitOfMeasure");

    const subUslXml = `<Field Type='Number' Name='SubUSL' StaticName='SubUSL' DisplayName='SubUSL' />`;
    await ensureField(subUslXml, "SubUSL");

    const unitOfMeasureXml = `<Field Type='Text' Name='UnitOfMeasure' StaticName='UnitOfMeasure' DisplayName='UnitOfMeasure' MaxLength='255' />`;
    await ensureField(unitOfMeasureXml, "UnitOfMeasure");

    const uslXml = `<Field Type='Number' Name='USL' StaticName='USL' DisplayName='USL' />`;
    await ensureField(uslXml, "USL");

    // Ensure some fields are added to the default view
    const view = list.defaultView;
    const schemaXml = await view.fields.getSchemaXml();
    const fieldsToEnsureInView = [
        "LinkTitle",
        "DataSource",
        "AssociatedPPM",
        "BaselineAndRevisionFrequency",
        "BG",
        "CausalAnalysisTrigger",
        "DataAnalysisFrequency",
        "DataCollectionFrequency",
        "Goal",
        "HasSubProcess",
        "InterpretationGuidelines",
        "USL",
        "LSL",
        "Metrics",
        "MetricsFormulae",
        "OrgCausalAnalysisTrigger",
        "OrgInterpretationGuidelines",
        "OrgStatistical",
        "PerformanceGoals",
        "PG",
        "Priority",
        "ProbabilityOfSuccessThreshold",
        "Process",
        "ProjectType",
        "Quantitative",
        "UnitOfMeasure",
        "Statistical",
        "SubApplicability",
        "SubBaselineAndRevisionFrequency",
        "SubDataAnalysisFrequency",
        "SubDataCollectionFrequency",
        "SubDataInput",
        "SubDataSource",
        "SubGoal",
        "SubLSL",
        "SubMetrics",
        "SubMetricsFormulae",
        "Subprocess",
        "SubUnitOfMeasure",
        "SubUSL",
        "DataInput",
        "VersionId",
        "Modified",
        "Created"
    ];
    for (const f of fieldsToEnsureInView) {
        if (!schemaXml.includes(`Name=\"${f}\"`) && !schemaXml.includes(`Name='${f}'`)) {
            await view.fields.add(f);
        }
    }

}

export default provisionProjectMetrics;