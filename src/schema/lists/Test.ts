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

	// Ensure each site content type exists and add it to the list.
	for (const ctId in CONTENT_TYPES) {
		if (!Object.prototype.hasOwnProperty.call(CONTENT_TYPES, ctId)) {
			continue;
		}
		const ctName = CONTENT_TYPES[ctId];

		if (!await contentTypeExists(ctId)) {
			try {
				// Use the older, widely-supported signature: add(name, id, group)
				await (sp.web.contentTypes as any).add(ctName, ctId, 'Custom');
			} catch (e) {
				// If creation fails, continue - adding to list may still work if the CT exists elsewhere
			}
		}

		// Try to add the content type to the list. Be flexible with available method names across PnP versions.
		try {
			const listCt = (list && (list.contentTypes || (sp.web.lists.getByTitle(LIST_TITLE).contentTypes))) as any;
			if (listCt) {
				if (typeof listCt.addAvailableContentType === 'function') {
					await listCt.addAvailableContentType(ctId);
				} else if (typeof listCt.addExistingContentType === 'function') {
					await listCt.addExistingContentType(ctId);
				} else if (typeof listCt.add === 'function') {
					// some versions accept an object or id
					try {
						await listCt.add({ Id: { StringValue: ctId }, Name: ctName } as any);
					} catch (e) {
						// ignore
					}
				} else {
					// last resort: call through the web lists path
					await sp.web.lists.getByTitle(LIST_TITLE).contentTypes.addAvailableContentType(ctId);
				}
			}
		} catch (e) {
			// ignore errors here — provisioning should continue.
		}
	}

	// After adding our custom content types, set them as the list's content type order
	// and remove the default Item content type so the New menu shows only our CTs.
	try {
		const customCtIds: string[] = [];
		for (const k in CONTENT_TYPES) {
			if (Object.prototype.hasOwnProperty.call(CONTENT_TYPES, k)) {
				customCtIds.push(k);
			}
		}

		// Set the content type order on the list's root folder so the first custom CT becomes the default/new form
		try {
			const rootFolder = (sp.web.lists.getByTitle(LIST_TITLE) as any).rootFolder as any;
			if (customCtIds.length > 0) {
				await rootFolder.update({ ContentTypeOrder: customCtIds.map(id => ({ StringValue: id })) });
			}
		} catch (e) {
			// ignore failures to update root folder order
		}

		// Attempt to remove the out-of-the-box Item content type (id 0x01) from the list so it doesn't appear in New menu
		try {
			// Retrieve the content types currently attached to the list and log/return them.
			// Use a select to get Id and Name; cast to any to be compatible with repo PnPJS typings.
			const cts = await (sp.web.lists.getByTitle(LIST_TITLE) as any).contentTypes.select("Id", "Name")();
			// Example: log them for debugging during provisioning
			// Find the content type entry that corresponds to the built-in "Item" content type
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
					// Match by name 'Item' or id starting with 0x01 (built-in Item CT)
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
			} else {
				// No Item CT found in list's content types
				console.log('No Item content type found on list', LIST_TITLE, cts);
			}
		} catch (e) {
			// ignore - reading content types may fail in some PnPJS versions/environments
		}
	} catch (e) {
		// ignore top-level errors
	}

	// Ensure some basic fields are in the default view
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

/**
 * Helper: return the content types attached to a list (Id and Name)
 */
export async function getListContentTypes(sp: SPFI, listTitle: string = LIST_TITLE): Promise<any[]> {
	try {
		const cts = await (sp.web.lists.getByTitle(listTitle) as any).contentTypes.select("Id", "Name")();
		return cts || [];
	} catch (e) {
		return [];
	}
}

