import { SPFI } from "@pnp/sp";
import { RequiredLists } from '../RequiredList';
// ensure module augmentation for webs/lists is loaded so SPFI has `.web` and lists types
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/views";

const LIST_TITLE = RequiredLists.EmailLogs;

export async function provisionEmailLogs(sp: SPFI): Promise<void> {
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
        const ensureResult = await sp.web.lists.ensure(LIST_TITLE, "Email logs list", 100);
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

    const statusXml = `<Field Type='Text' Name='Status' StaticName='Status' DisplayName='Status' MaxLength='255' />`;
    await ensureField(statusXml, "Status");

    const mailSentToXml = `<Field Type='Text' Name='MailSentTo' StaticName='MailSentTo' DisplayName='MailSentTo' MaxLength='255' />`;
    await ensureField(mailSentToXml, "MailSentTo");

    const versionIdXml = `<Field Type='Text' Name='VersionId' StaticName='VersionId' DisplayName='VersionId' MaxLength='255' />`;
    await ensureField(versionIdXml, "VersionId");

    const mailSentStatusXml = `<Field Type='Note' Name='MailSentStatus' StaticName='MailSentStatus' DisplayName='MailSentStatus' NumLines='6' RichText='FALSE' />`;
    await ensureField(mailSentStatusXml, "MailSentStatus");

    const view = list.defaultView;
    const schemaXml = await view.fields.getSchemaXml();
    const fieldsToEnsureInView = ["LinkTitle", "Status", "MailSentTo", "VersionId", "MailSentStatus", "Author", "Created"];
    for (const f of fieldsToEnsureInView) {
        if (!schemaXml.includes(`Name=\"${f}\"`) && !schemaXml.includes(`Name='${f}'`)) {
            await view.fields.add(f);
        }
    }
}

export default provisionEmailLogs;