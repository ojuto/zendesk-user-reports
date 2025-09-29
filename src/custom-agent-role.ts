export interface CustomAgentRole {
    name: string;
    description: string;
    id: number;
    created_at: string;
    role_type: number;
    team_member_count: number;
    updatedAt: string;
}

export interface CustomAgentRoleResponse {
    custom_roles: CustomAgentRole[];
    meta?: { has_more?: boolean };
    links?: { next?: string };
}
