export interface Brand {
    name: string;
    id: number;
    created_at: string;
    updatedAt: string;
}

export interface BrandResponse {
    brands: Brand[];
    meta?: { has_more?: boolean };
    links?: { next?: string };
}

export function filterVorwerkInternational(brands: Brand[]): Brand[] {
    return brands.filter((brand) => brand.name !== 'Vorwerk International');
}
