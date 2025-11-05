import { SPFI } from "@pnp/sp";
import { RequiredLists } from '../RequiredList';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";

const LIST_TITLE = RequiredLists.ProjectMetricLogs;

export async function provisionProjectMetricLogs(sp: SPFI): Promise<void> {
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
        const ensureResult = await sp.web.lists.ensure(LIST_TITLE, "Project metrics logs list", 100);
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

    const versionIdXml = `<Field Type='Text' Name='VersionId' StaticName='VersionId' DisplayName='VersionId' MaxLength='255' />`;
    await ensureField(versionIdXml, "VersionId");

    const statusChoices = ["Draft", "In Review", "In Approval", "Approved", "Rejected"];
    const statusChoicesXml = statusChoices.map(c => `<CHOICE>${c}</CHOICE>`).join("");
    const statusXml = `<Field Type='Choice' Name='Status' StaticName='Status' DisplayName='Status' Format='Dropdown'><CHOICES>${statusChoicesXml}</CHOICES></Field>`;
    await ensureField(statusXml, "Status");

    const pmCommentsXml = `<Field Type='Note' Name='PMComments' StaticName='PMComments' DisplayName='PMComments' NumLines='6' RichText='FALSE' />`;
    await ensureField(pmCommentsXml, "PMComments");

    const reviewerCommentsXml = `<Field Type='Note' Name='ReviewerComments' StaticName='ReviewerComments' DisplayName='ReviewerComments' NumLines='6' RichText='FALSE' />`;
    await ensureField(reviewerCommentsXml, "ReviewerComments");

    const createdVersionChoices = ["Minor", "Major"];
    const createdVersionChoicesXml = createdVersionChoices.map(c => `<CHOICE>${c}</CHOICE>`).join("");
    const createdVersionXml = `<Field Type='Choice' Name='CreatedVersion' StaticName='CreatedVersion' DisplayName='CreatedVersion' Format='Dropdown'><CHOICES>${createdVersionChoicesXml}</CHOICES></Field>`;
    await ensureField(createdVersionXml, "CreatedVersion");

    const isActiveXml = `<Field Type='Boolean' Name='IsActive' StaticName='IsActive' DisplayName='IsActive' />`;
    await ensureField(isActiveXml, "IsActive");

    const view = list.defaultView;
    const schemaXml = await view.fields.getSchemaXml();
    const fieldsToEnsureInView = ["VersionId", "Status", "PMComments", "ReviewerComments", "CreatedVersion", "IsActive"];
    for (const f of fieldsToEnsureInView) {
        if (!schemaXml.includes(`Name=\"${f}\"`) && !schemaXml.includes(`Name='${f}'`)) {
            await view.fields.add(f);
        }
    }

}

export default provisionProjectMetricLogs;