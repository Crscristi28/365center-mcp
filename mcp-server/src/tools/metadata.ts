import { graphClient, callSharePointRest } from "../auth.js";

// Translate drive item ID to list item ID if needed
async function resolveListItemId(siteId: string, listId: string, itemId: string, driveId?: string): Promise<string> {
  // If it's already a numeric ID, use it directly
  if (/^\d+$/.test(itemId)) {
    return itemId;
  }

  // It's a drive item ID — resolve to list item ID
  if (!driveId) {
    throw new Error("driveId is required when using a drive item ID. Use list_document_libraries to get the drive ID.");
  }

  const result = await graphClient
    .api(`/drives/${driveId}/items/${itemId}/listItem`)
    .select("id")
    .get();

  return result.id;
}

export async function listColumns(siteId: string, listId: string) {
  const result = await graphClient
    .api(`/sites/${siteId}/lists/${listId}/columns`)
    .get();

  return result.value
    .filter((col: any) => !col.readOnly)
    .map((col: any) => ({
      id: col.id,
      name: col.name,
      displayName: col.displayName,
      type: col.text ? "text" :
            col.choice ? "choice" :
            col.boolean ? "boolean" :
            col.dateTime ? "dateTime" :
            col.number ? "number" :
            col.lookup ? "lookup" :
            col.personOrGroup ? "personOrGroup" :
            "other",
      required: col.required,
      choices: col.choice?.choices,
      allowMultipleValues: col.choice?.allowMultipleValues,
    }));
}

export async function createChoiceColumn(
  siteId: string,
  listId: string,
  name: string,
  displayName: string,
  choices: string[],
  allowMultiple: boolean = false
) {
  if (!allowMultiple) {
    const columnDef = {
      name,
      displayName,
      choice: {
        allowTextEntry: false,
        choices,
        displayAs: "dropDownMenu" as const,
      },
    };

    const result = await graphClient
      .api(`/sites/${siteId}/lists/${listId}/columns`)
      .post(columnDef);

    return {
      id: result.id,
      name: result.name,
      displayName: result.displayName,
      allowMultiple: false,
    };
  }

  const site = await graphClient.api(`/sites/${siteId}`).select("webUrl").get();
  const siteUrl = site.webUrl;

  const listIdentifier = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(listId)
    ? `lists(guid'${listId}')`
    : `lists/getByTitle('${listId}')`;

  const result = await callSharePointRest(
    siteUrl,
    `/_api/web/${listIdentifier}/fields`,
    "POST",
    {
      __metadata: { type: "SP.FieldMultiChoice" },
      FieldTypeKind: 15,
      Title: displayName,
      StaticName: name,
      InternalName: name,
      Choices: { results: choices },
    }
  ) as any;

  return {
    id: result.d?.Id || result.Id,
    name: name,
    displayName: displayName,
    allowMultiple: true,
  };
}

export async function deleteColumn(siteId: string, listId: string, columnId: string) {
  await graphClient
    .api(`/sites/${siteId}/lists/${listId}/columns/${columnId}`)
    .delete();
  return { success: true, columnId };
}

export async function createTextColumn(
  siteId: string,
  listId: string,
  name: string,
  displayName: string
) {
  const result = await graphClient
    .api(`/sites/${siteId}/lists/${listId}/columns`)
    .post({
      name,
      displayName,
      text: {
        allowMultipleLines: false,
        maxLength: 255,
      },
    });

  return {
    id: result.id,
    name: result.name,
    displayName: result.displayName,
  };
}

export async function setDocumentMetadata(
  siteId: string,
  listId: string,
  itemId: string,
  fields: Record<string, unknown>,
  driveId?: string
) {
  const resolvedId = await resolveListItemId(siteId, listId, itemId, driveId);

  const patchFields: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(fields)) {
    if (Array.isArray(value)) {
      patchFields[`${key}@odata.type`] = "Collection(Edm.String)";
      patchFields[key] = value;
    } else {
      patchFields[key] = value;
    }
  }

  const result = await graphClient
    .api(`/sites/${siteId}/lists/${listId}/items/${resolvedId}/fields`)
    .patch(patchFields);

  return result;
}

export async function getDocumentMetadata(
  siteId: string,
  listId: string,
  itemId: string,
  driveId?: string
) {
  const resolvedId = await resolveListItemId(siteId, listId, itemId, driveId);

  const result = await graphClient
    .api(`/sites/${siteId}/lists/${listId}/items/${resolvedId}/fields`)
    .get();

  return result;
}
