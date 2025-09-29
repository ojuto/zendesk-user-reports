import axios from 'axios';

import { requireEnv } from './utils';

import type { Brand, BrandResponse } from './brand';
import type { BrandAgent, BrandAgentResponse } from './brand-agent';
import type { CustomAgentRole, CustomAgentRoleResponse } from './custom-agent-role';
import type { UserResponse, User } from './user';
import type { AxiosResponse } from 'axios';

export const authVi = Buffer.from(
    `${requireEnv('ZENDESK_VI_EMAIL')}/token:${requireEnv('ZENDESK_VI_API_TOKEN')}`,
).toString('base64');

export const authVde = Buffer.from(
    `${requireEnv('ZENDESK_VDE_EMAIL')}/token:${requireEnv('ZENDESK_VDE_API_TOKEN')}`,
).toString('base64');

export const baseUrlVi = requireEnv('ZENDESK_VI_BASE_URL');
export const baseUrlVde = requireEnv('ZENDESK_VDE_BASE_URL');

export async function fetchAllUsers(baseUrl: string, auth: string): Promise<User[]> {
    const users: User[] = [];
    let url: string | null = `${baseUrl}/api/v2/users.json`;

    while (url) {
        const res: AxiosResponse<UserResponse> = await axios.get(url, {
            headers: { Authorization: `Basic ${auth}` },
            params: {
                'page[size]': '100',
                'role[]': ['admin', 'agent'],
            },
        });

        users.push(...res.data.users);

        console.log(`${users.length} users fetched from ${baseUrl}`);

        if (!res.data.meta?.has_more) {
            break;
        }

        url = res.data.links?.next ?? null;
    }

    return users;
}

export async function fetchAllCustomAgentRoles(
    baseUrl: string,
    auth: string,
): Promise<CustomAgentRole[]> {
    const roles: CustomAgentRole[] = [];
    let url: string | null = `${baseUrl}/api/v2/custom_roles.json`;

    while (url) {
        const res: AxiosResponse<CustomAgentRoleResponse> = await axios.get(url, {
            headers: { Authorization: `Basic ${auth}` },
        });

        roles.push(...res.data.custom_roles);

        console.log(`${roles.length} roles fetched from ${baseUrl}`);

        if (!res.data.meta?.has_more) {
            break;
        }

        url = res.data.links?.next ?? null;
    }

    return roles;
}

export async function fetchAllBrands(baseUrl: string, auth: string): Promise<Brand[]> {
    const brands: Brand[] = [];
    let url: string | null = `${baseUrl}/api/v2/brands.json`;

    while (url) {
        const res: AxiosResponse<BrandResponse> = await axios.get(url, {
            headers: { Authorization: `Basic ${auth}` },
            params: {
                'page[size]': '100',
            },
        });

        brands.push(...res.data.brands);

        console.log(`${brands.length} brands fetched from ${baseUrl}`);

        if (!res.data.meta?.has_more) {
            break;
        }

        url = res.data.links?.next ?? null;
    }

    return brands;
}

export async function fetchAllBrandAgents(baseUrl: string, auth: string): Promise<BrandAgent[]> {
    const brandAgents: BrandAgent[] = [];
    let url: string | null = `${baseUrl}/api/v2/brand_agents.json`;

    while (url) {
        const res: AxiosResponse<BrandAgentResponse> = await axios.get(url, {
            headers: { Authorization: `Basic ${auth}` },
            params: {
                'page[size]': '100',
            },
        });

        brandAgents.push(...res.data.brand_agents);

        console.log(`${brandAgents.length} brand agents fetched from ${baseUrl}`);

        if (!res.data.meta?.has_more) {
            break;
        }

        url = res.data.links?.next ?? null;
    }

    return brandAgents;
}
