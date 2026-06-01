/**
 * People and directory lookup API.
 * Uses OWA's PeopleGraphVx endpoint for profile lookups.
 */

import { getOwaToken } from '../auth/index.js';
import { OWA_PEOPLE, OWA_BASE } from '../constants.js';
import { getBearerHeaders, parseResponse, fetchWithRetry } from '../utils/http.js';
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

// ─────────────────────────────────────────────────────────────────────────────
// Get person by email
// ─────────────────────────────────────────────────────────────────────────────

export async function getPersonByEmail(email: string): Promise<{ displayName?: string; photoUrl?: string } | null> {
  const token = await getOwaToken();
  if (!token) return null;

  const url = `${OWA_PEOPLE}/people/SMTP:${email}/fetchResizedPhoto(height=120,width=120,allowResizeUp=true)/$value`;

  try {
    const res = await fetchWithRetry(url, {
      method: 'GET',
      headers: getBearerHeaders(token, OWA_BASE),
    });
    if (res.ok) {
      return { photoUrl: url };
    }
    return null;
  } catch {
    return null;
  }
}
