import type { CustomAgentRole } from './custom-agent-role';

export interface User {
    id: number;
    name: string;
    email: string | null;
    role: 'end-user' | 'agent' | 'admin';
    created_at: string;
    last_login_at: string | null;
    tags?: string[];
    details?: string;
    notes?: string;
    suspended?: boolean;
    active?: boolean;
    user_fields?: Record<string, unknown>;
    custom_role_id?: number;
    role_type: number;
}

export interface UserWithCustomAgentRole extends User {
    custom_agent_role_name?: string;
}

export interface UserResponse {
    users: User[];
    meta?: { has_more?: boolean };
    links?: { next?: string };
}

export function filterAgentsInactive45Days(users: User[]): User[] {
    const now = new Date();
    const msInDay = 1000 * 60 * 60 * 24;

    return users.filter((user) => {
        if (user.role !== 'agent') {
            return false;
        }

        if (user.suspended) {
            return false;
        }

        if (user.role_type === 1) {
            return false;
        }

        const tags = user.tags ?? [];
        if (tags.includes('lightagent_active') || tags.includes('cc_service_addresses')) {
            return false;
        }

        const lastLogin = new Date(user.last_login_at);
        const diffDays = (now.getTime() - lastLogin.getTime()) / msInDay;
        return diffDays >= 45;
    });
}

export function filterSuspendedAgentsNotLightAgents(users: User[]): User[] {
    return users.filter((user) => {
        if (user.role !== 'agent') {
            return false;
        }

        if (!user.suspended) {
            return false;
        }

        return user.role_type !== 1;
    });
}

export function filterLightAgentActiveButNotLightAgent(users: User[]): User[] {
    return users.filter((user) => {
        if (user.role !== 'agent') {
            return false;
        }

        if (user.role_type === 1) {
            return false;
        }

        const tags = user.tags ?? [];
        return tags.includes('lightagent_active');
    });
}

export function filterFunctionalUsers(users: User[]): User[] {
    return users.filter((user) => {
        const tags = user.tags ?? [];
        return tags.includes('functional_user');
    });
}

export function filterBrandRoleCount(users: User[]): User[] {
    return users.filter((user) => !user.suspended);
}

export function filterCommonUsers(users: User[]): User[] {
    return users.filter((user) => {
        if (user.suspended) {
            return false;
        }

        return user.role_type !== 1;
    });
}

export function countUsedSeats(usersVi: User[], usersVde: User[]): number {
    const usedSeatsVi = usersVi.filter((user) => {
        if (user.suspended) {
            return false;
        }

        return user.role_type !== 1;
    }).length;
    const usedSeatsVde = usersVde.filter((user) => {
        if (user.suspended) {
            return false;
        }

        return user.role_type !== 1;
    }).length;
    const totalLength = usedSeatsVi + usedSeatsVde;
    return totalLength - 7;
}

export function mapUsersWithCustomAgentRoleName(
    users: User[],
    customRoles: CustomAgentRole[],
): UserWithCustomAgentRole[] {
    const roleNameById = new Map<number, string>(customRoles.map((r) => [r.id, r.name]));

    return users.map((user) => {
        const roleName =
            user.custom_role_id != null ? roleNameById.get(user.custom_role_id) : undefined;

        return {
            ...user,
            custom_agent_role_name: roleName,
        };
    });
}
