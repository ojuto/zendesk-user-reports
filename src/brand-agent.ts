export interface BrandAgent {
    brand_id: number;
    user_id: number;
    id: number;
    created_at: string;
    updatedAt: string;
}

export interface BrandAgentResponse {
    brand_agents: BrandAgent[];
    meta?: { has_more?: boolean };
    links?: { next?: string };
}
