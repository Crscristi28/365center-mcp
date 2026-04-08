import { graphClient } from "../auth.js";

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
  const columnDef: any = {
    name,
    displayName,
    choice: {
      allowTextEntry: false,
      choices,
      displayAs: allowMultiple ? "checkBoxes" : "dropDownMenu",
    },
  };

  if (allowMultiple) {
    columnDef.indexed = false;
  }

  const result = await graphClient
    .api(`/sites/${siteId}/lists/${listId}/columns`)
    .post(columnDef);

  return {
    id: result.id,
    name: result.name,
    displayName: result.displayName,
    allowMultiple,
  };
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

  const result = await graphClient
    .api(`/sites/${siteId}/lists/${listId}/items/${resolvedId}/fields`)
    .patch(fields);

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
