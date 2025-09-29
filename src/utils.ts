export function formatDate(dateStr?: string | null): string {
    if (!dateStr) {
        return '';
    }
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) {
        return '';
    }
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    return `${day}.${month}.${year} ${hours}:${minutes}`;
}

export function daysSince(dateStr?: string | null): string {
    if (!dateStr) {
        return '-';
    }
    const lastLogin = new Date(dateStr);
    if (isNaN(lastLogin.getTime())) {
        return '-';
    }
    const now = new Date();
    const diffMs = now.getTime() - lastLogin.getTime();
    const diffDays = diffMs / (1000 * 60 * 60 * 24);
    return diffDays.toFixed(1);
}

export const requireEnv = (name: string): string => {
    const environmentVariable = process.env[name];
    if (!environmentVariable) {
        throw new Error(`Missing required environment variable: ${name}`);
    }
    return environmentVariable;
};
