/**
 * People and directory lookup API.
 */

import { owaGet } from './client.js';
import type { ODataResponse } from './mail.js';

// ─────────────────────────────────────────────────────────────────────────────
// Types
// ─────────────────────────────────────────────────────────────────────────────

export interface Person {
  Id?: string;
  DisplayName: string;
  GivenName?: string;
  Surname?: string;
  EmailAddresses?: Array<{ Name?: string; Address: string }>;
  Phones?: Array<{ Number: string; Type: string }>;
  JobTitle?: string;
  Department?: string;
  OfficeLocation?: string;
  CompanyName?: string;
  UserPrincipalName?: string;
  ScoredEmailAddresses?: Array<{ Address: string; RelevanceScore: number }>;
}

// ─────────────────────────────────────────────────────────────────────────────
// Search people
// ─────────────────────────────────────────────────────────────────────────────

export async function searchPeople(query: string, top = 10): Promise<Person[]> {
  const res = await owaGet<ODataResponse<Person>>('/people', {
    '$search': `"${query}"`,
    '$top': String(top),
    '$select': 'DisplayName,GivenName,Surname,ScoredEmailAddresses,JobTitle,Department,UserPrincipalName',
  });
  return res.value;
}
