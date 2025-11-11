import { SPFI } from "@pnp/sp";
import { RequiredLists } from '../RequiredList';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/content-types";
import "@pnp/sp/fields";
import "@pnp/sp/views";

const LIST_TITLE = RequiredLists.Test;
const CONTENT_TYPES: { [id: string]: string } = {
	'0x0100675A5D902917F5468E54C67BEC4A6765': 'Lessons Learnt',
	'0x010013D4E57092D07541AB01F187F1C1A283': 'Best Practices',
	'0x010063DA7C6C73EA594FA193E1333337F0E1': 'Reusable Component'
};

export async function provisionTest(sp: SPFI): Promise<void> {
	async function listExists(title: string): Promise<boolean> {
		try {
			await sp.web.lists.getByTitle(title).select("Id")();
			return true;
		} catch (e) {
			return false;
		}
	}


	async function contentTypeExists(id: string): Promise<boolean> {
		try {
			await sp.web.contentTypes.getById(id).select("Id")();
			return true;
		} catch (e) {
			return false;
		}
	}

	let list: any;
	const exists = await listExists(LIST_TITLE);
	if (!exists) {
		const ensureResult = await sp.web.lists.ensure(LIST_TITLE, "Test list", 100, true);
		list = ensureResult.list;
	} else {
		list = sp.web.lists.getByTitle(LIST_TITLE);
		try {
			await list.update({ ContentTypesEnabled: true });
		} catch (e) {
			// ignore if update not supported
		}
	}

	for (const ctId in CONTENT_TYPES) {
		if (!Object.prototype.hasOwnProperty.call(CONTENT_TYPES, ctId)) {
			continue;
		}
		const ctName = CONTENT_TYPES[ctId];

		if (!await contentTypeExists(ctId)) {
			try {
				await (sp.web.contentTypes as any).add(ctName, ctId, 'Custom');
			} catch (e) {
				// If creation fails, continue - adding to list may still work if the CT exists elsewhere
			}
		}

		try {
			const listCt = (list && (list.contentTypes || (sp.web.lists.getByTitle(LIST_TITLE).contentTypes))) as any;
			if (listCt) {
				if (typeof listCt.addAvailableContentType === 'function') {
					await listCt.addAvailableContentType(ctId);
				} else if (typeof listCt.addExistingContentType === 'function') {
					await listCt.addExistingContentType(ctId);
				} else if (typeof listCt.add === 'function') {
					try {
						await listCt.add({ Id: { StringValue: ctId }, Name: ctName } as any);
					} catch (e) {
						// ignore
					}
				} else {
					await sp.web.lists.getByTitle(LIST_TITLE).contentTypes.addAvailableContentType(ctId);
				}
			}
		} catch (e) {
			// ignore errors here — provisioning should continue.
		}
	}

	try {
		const customCtIds: string[] = [];
		for (const k in CONTENT_TYPES) {
			if (Object.prototype.hasOwnProperty.call(CONTENT_TYPES, k)) {
				customCtIds.push(k);
			}
		}

		try {
			const rootFolder = (sp.web.lists.getByTitle(LIST_TITLE) as any).rootFolder as any;
			if (customCtIds.length > 0) {
				await rootFolder.update({ ContentTypeOrder: customCtIds.map(id => ({ StringValue: id })) });
			}
		} catch (e) {
			// ignore failures to update root folder order
		}

		try {
			const cts = await (sp.web.lists.getByTitle(LIST_TITLE) as any).contentTypes.select("Id", "Name")();
			const unwrapId = (idField: any): string => {
				if (!idField) { return ''; }
				if (typeof idField === 'string') { return idField; }
				if (typeof idField === 'object' && idField.StringValue) { return idField.StringValue; }
				return '';
			};

			const itemCt = (cts || []).find((ct: any) => {
				try {
					const name = (ct && ct.Name) ? String(ct.Name) : '';
					const idVal = unwrapId(ct && ct.Id).toLowerCase();
					return name === 'Item' || idVal.indexOf('0x01') === 0;
				} catch (e) {
					return false;
				}
			});

			if (itemCt) {
				const itemId = unwrapId(itemCt.Id);
				if (itemId) {
					try {
						await (sp.web.lists.getByTitle(LIST_TITLE) as any).contentTypes.getById(itemId).delete();
						console.log(`Deleted Item content type (${itemId}) from list ${LIST_TITLE}`);
					} catch (e) {
						// ignore deletion errors (in use / locked / permission)
					}
				}
			}
		} catch (e) {
			// ignore - reading content types may fail in some PnPJS versions/environments
		}
	} catch (e) {
		// ignore top-level errors
	}

	try {
		const view = list.defaultView;
		const schemaXml = await view.fields.getSchemaXml();
		const fieldsToEnsureInView = ["LinkTitle", "ContentType", "Author", "Created"];
		for (const f of fieldsToEnsureInView) {
			if (!schemaXml.includes(`Name=\"${f}\"`) && !schemaXml.includes(`Name='${f}'`)) {
				await view.fields.add(f);
			}
		}
	} catch (e) {
		// ignore view errors
	}
}

export default provisionTest;

export async function getListContentTypes(sp: SPFI, listTitle: string = LIST_TITLE): Promise<any[]> {
	try {
		const cts = await (sp.web.lists.getByTitle(listTitle) as any).contentTypes.select("Id", "Name")();
		return cts || [];
	} catch (e) {
		return [];
	}
}

