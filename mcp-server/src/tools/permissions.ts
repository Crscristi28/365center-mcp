import { callSharePointRest } from "../auth.js";

export async function getSitePermissions(siteUrl: string) {
  // Get all SharePoint groups for this site
  const result = await callSharePointRest(siteUrl, "/_api/web/sitegroups", "GET") as any;

  return result.d.results.map((group: any) => ({
    id: group.Id,
    title: group.Title,
    description: group.Description,
    ownerTitle: group.OwnerTitle,
    userCount: group.Users ? group.Users.results?.length : undefined,
  }));
}

export async function getGroupMembers(siteUrl: string, groupId: number) {
  const result = await callSharePointRest(
    siteUrl,
    `/_api/web/sitegroups/getbyid(${groupId})/users`,
    "GET"
  ) as any;

  return result.d.results.map((user: any) => ({
    id: user.Id,
    title: user.Title,
    email: user.Email,
    loginName: user.LoginName,
  }));
}

export async function addUserToGroup(siteUrl: string, groupId: number, userEmail: string) {
  const result = await callSharePointRest(
    siteUrl,
    `/_api/web/sitegroups/getbyid(${groupId})/users`,
    "POST",
    {
      "__metadata": { "type": "SP.User" },
      "LoginName": `i:0#.f|membership|${userEmail}`,
    }
  ) as any;

  return {
    id: result.d.Id,
    title: result.d.Title,
    email: result.d.Email,
    groupId,
  };
}

export async function removeUserFromGroup(siteUrl: string, groupId: number, userId: number) {
  await callSharePointRest(
    siteUrl,
    `/_api/web/sitegroups/getbyid(${groupId})/users/removebyid(${userId})`,
    "POST"
  );

  return { success: true, userId, groupId };
}
