import { callSharePointRest } from "../auth.js";

export async function getSitePermissions(siteUrl: string, includeMembers: boolean = false) {
  // Default: groups only. With includeMembers: expand Users and select minimal fields
  // to avoid the N+1 pattern (get_permissions + N × get_group_members).
  const apiPath = includeMembers
    ? "/_api/web/sitegroups?$expand=Users&$select=Id,Title,Description,OwnerTitle,Users/Id,Users/Title,Users/Email,Users/LoginName"
    : "/_api/web/sitegroups";

  const result = await callSharePointRest(siteUrl, apiPath, "GET") as any;

  return result.d.results.map((group: any) => {
    const base = {
      id: group.Id,
      title: group.Title,
      description: group.Description,
      ownerTitle: group.OwnerTitle,
    };
    if (!includeMembers) return base;
    return {
      ...base,
      members: (group.Users?.results || []).map((u: any) => ({
        id: u.Id,
        title: u.Title,
        email: u.Email,
        loginName: u.LoginName,
      })),
    };
  });
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
