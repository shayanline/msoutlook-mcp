/**
 * Directory API: org hierarchy and rich user profiles via Microsoft Graph.
 *
 * Works with the directory read permission the Outlook app already has. Pass an
 * email/UPN to target a colleague, or omit it to target the signed-in user.
 */

import { graphGetPath, graphGetBinary } from './client.js';

export interface DirectoryUser {
  id?: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  mail?: string;
  userPrincipalName?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  mobilePhone?: string;
  businessPhones?: string[];
  city?: string;
  country?: string;
}

const PROFILE_SELECT =
  'id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones,city,country';

/** Build the base Graph path for a target: '/me' or '/users/{email}'. */
function base(email?: string): string {
  return email ? `/users/${encodeURIComponent(email)}` : '/me';
}

export async function getUserProfile(email?: string): Promise<DirectoryUser> {
  return graphGetPath<DirectoryUser>(base(email), { '$select': PROFILE_SELECT });
}

export async function getManager(email?: string): Promise<DirectoryUser> {
  return graphGetPath<DirectoryUser>(`${base(email)}/manager`, { '$select': PROFILE_SELECT });
}

export async function getDirectReports(email?: string): Promise<DirectoryUser[]> {
  const res = await graphGetPath<{ value: DirectoryUser[] }>(`${base(email)}/directReports`, {
    '$select': PROFILE_SELECT,
  });
  return res.value ?? [];
}

export async function getUserPhoto(email?: string): Promise<{ contentType: string; bytes: Buffer }> {
  return graphGetBinary(`${base(email)}/photo/$value`);
}
