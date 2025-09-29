import { daysSince, formatDate } from './utils';

import type { Brand } from './brand';
import type { BrandAgent } from './brand-agent';
import type { CustomAgentRole } from './custom-agent-role';
import type { User, UserWithCustomAgentRole } from './user';
import type Worksheet from 'exceljs/index';

export function generateUserWorksheet(
    sheet: Worksheet,
    usersVi: UserWithCustomAgentRole[],
    usersVde: UserWithCustomAgentRole[],
): void {
    sheet.getRow(2).height = 18;

    const columns: string[] = [
        'Name',
        'Mail',
        'Role',
        'Created at',
        'Details',
        'Notes',
        'Tags',
        'Last login in days',
    ];

    const offsetVi = 1;
    sheet.mergeCells(1, offsetVi, 1, offsetVi + columns.length - 1);
    const viHeader = sheet.getCell(1, offsetVi);
    viHeader.value = 'VI';
    viHeader.alignment = { horizontal: 'center', vertical: 'middle' };
    viHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    viHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF03363D' } };

    const headerPadding = 3;
    columns.forEach((col, i) => {
        const column = sheet.getColumn(offsetVi + i);
        column.width = Math.max(col.length + headerPadding, 12);
    });

    const offsetVde = columns.length + 3;
    sheet.mergeCells(1, offsetVde, 1, offsetVde + columns.length - 1);
    const vdeHeader = sheet.getCell(1, offsetVde);
    vdeHeader.value = 'VDE';
    vdeHeader.alignment = { horizontal: 'center', vertical: 'middle' };
    vdeHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    vdeHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF03363D' } };

    columns.forEach((col, i) => {
        const column = sheet.getColumn(offsetVde + i);
        column.width = Math.max(col.length + headerPadding, 12);
    });

    const safeSheetName = sheet.name.replace(/[^A-Za-z0-9_]/g, '_');
    const tableNameVI = `TableVI_${safeSheetName}`;
    const tableNameVDE = `TableVDE_${safeSheetName}`;

    const rowsVi = usersVi.map((user) => [
        user.name,
        user.email ?? '',
        user.custom_agent_role_name,
        formatDate(user.created_at),
        user.details,
        user.notes,
        (user.tags ?? []).join(', '),
        daysSince(user.last_login_at),
    ]);

    const rowsVde = usersVde.map((user) => [
        user.name,
        user.email ?? '',
        user.custom_agent_role_name,
        formatDate(user.created_at),
        user.details,
        user.notes,
        (user.tags ?? []).join(', '),
        daysSince(user.last_login_at),
    ]);

    const writeHeaderWithoutTable = (startCol: number) => {
        columns.forEach((name, i) => {
            const cell = sheet.getCell(2, startCol + i);
            cell.value = name;
            cell.font = { bold: true };
            cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        });
    };

    if (rowsVi.length > 0) {
        sheet.addTable({
            name: tableNameVI,
            ref: sheet.getCell(2, offsetVi).address,
            style: { theme: 'TableStyleLight1' },
            columns: columns.map((name) => ({ name, filterButton: true })),
            rows: rowsVi,
        });
    } else {
        writeHeaderWithoutTable(offsetVi);
    }

    if (rowsVde.length > 0) {
        sheet.addTable({
            name: tableNameVDE,
            ref: sheet.getCell(2, offsetVde).address,
            style: { theme: 'TableStyleLight1' },
            columns: columns.map((name) => ({ name, filterButton: true })),
            rows: rowsVde,
        });
    } else {
        writeHeaderWithoutTable(offsetVde);
    }

    const applyWrapAndBorder = (
        startRow: number,
        startCol: number,
        rowCount: number,
        colCount: number,
    ) => {
        const endRow = startRow + rowCount - 1;
        const endCol = startCol + colCount - 1;
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                const cell = sheet.getCell(r, c);
                const currentAlignment = cell.alignment ?? {};
                const isNumericCol = c === endCol;
                const desiredHorizontal: 'left' | 'right' = isNumericCol ? 'right' : 'left';
                cell.alignment = {
                    ...currentAlignment,
                    wrapText: true,
                    horizontal: desiredHorizontal,
                    vertical: 'top',
                };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
            }
        }
    };

    const dataStartRow = 3;
    if (rowsVi.length > 0) {
        applyWrapAndBorder(dataStartRow, offsetVi, rowsVi.length, columns.length);
    }
    if (rowsVde.length > 0) {
        applyWrapAndBorder(dataStartRow, offsetVde, rowsVde.length, columns.length);
    }
}

export function generateBrandRoleCountWorksheet(
    sheet: Worksheet,
    brandsVi: Brand[],
    brandsVde: Brand[],
    brandAgentsVi: BrandAgent[],
    brandAgentsVde: BrandAgent[],
    usersVi: User[],
    usersVde: User[],
    rolesVi: CustomAgentRole[],
    rolesVde: CustomAgentRole[],
): void {
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF03363D' } };
    const headerFont = { bold: true, color: { argb: 'FFFFFFFF' } };

    const columnsPerSide = 2;
    const gap = 3;
    const offsetVi = 1;
    const offsetVde = offsetVi + columnsPerSide + gap + 1;
    const offsetAgg = 4;

    sheet.mergeCells(1, offsetVi, 1, offsetVi + columnsPerSide - 1);
    const viHeader = sheet.getCell(1, offsetVi);
    viHeader.value = 'VI';
    viHeader.alignment = { horizontal: 'center', vertical: 'middle' };
    viHeader.font = headerFont;
    viHeader.fill = headerFill;

    sheet.mergeCells(1, offsetVde, 1, offsetVde + columnsPerSide - 1);
    const vdeHeader = sheet.getCell(1, offsetVde);
    vdeHeader.value = 'VDE';
    vdeHeader.alignment = { horizontal: 'center', vertical: 'middle' };
    vdeHeader.font = headerFont;
    vdeHeader.fill = headerFill;

    const pad = 4;
    const setColWidths = (startCol: number) => {
        sheet.getColumn(startCol).width = Math.max('Role'.length + pad, 18);
        sheet.getColumn(startCol + 1).width = Math.max('Team members'.length + pad, 14);
    };
    setColWidths(offsetVi);
    setColWidths(offsetVde);
    setColWidths(offsetAgg);

    type BrandRoleRows = Array<{ role: string; count: number }>;
    const buildCounts = (
        brands: Brand[],
        brandAgents: BrandAgent[],
        users: User[],
        roles: CustomAgentRole[],
    ): Map<number, { brandName: string; rows: BrandRoleRows }> => {
        const roleIdToName = new Map<number, string>(roles.map((r) => [r.id, r.name]));
        const userIdToRoleId = new Map<number, number | undefined>(
            users.map((u) => [u.id, (u as any).custom_role_id]),
        );
        const byBrand = new Map<number, { brandName: string; rows: BrandRoleRows }>();

        for (const brand of brands) {
            const agentsForBrand = brandAgents.filter((ba) => ba.brand_id === brand.id);
            const counts = new Map<string, number>();

            for (const ba of agentsForBrand) {
                const roleId = userIdToRoleId.get(ba.user_id);
                if (!roleId) continue;
                const roleName = roleIdToName.get(roleId);
                if (!roleName) continue;
                counts.set(roleName, (counts.get(roleName) ?? 0) + 1);
            }

            const rows: BrandRoleRows = Array.from(counts.entries())
                .map(([role, count]) => ({ role, count }))
                .sort((a, b) => b.count - a.count || a.role.localeCompare(b.role));

            byBrand.set(brand.id, { brandName: brand.name, rows });
        }

        return byBrand;
    };

    const viCounts = buildCounts(brandsVi, brandAgentsVi, usersVi, rolesVi);
    const vdeCounts = buildCounts(brandsVde, brandAgentsVde, usersVde, rolesVde);

    const writeSide = (
        startCol: number,
        brands: Brand[],
        counts: Map<number, { brandName: string; rows: BrandRoleRows }>,
    ): number => {
        let row = 2;

        const writeBrandBlock = (brandName: string, rows: BrandRoleRows) => {
            sheet.mergeCells(row, startCol, row, startCol + 1);
            const brandHeader = sheet.getCell(row, startCol);
            brandHeader.value = brandName;
            brandHeader.font = { bold: true };
            brandHeader.alignment = { horizontal: 'left', vertical: 'middle' };
            brandHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE6F0F2' } };
            row++;

            const roleHdr = sheet.getCell(row, startCol);
            roleHdr.value = 'Role';
            roleHdr.font = { bold: true };
            roleHdr.alignment = { horizontal: 'left', vertical: 'middle' };

            const cntHdr = sheet.getCell(row, startCol + 1);
            cntHdr.value = 'Team members';
            cntHdr.font = { bold: true };
            cntHdr.alignment = { horizontal: 'right', vertical: 'middle' };
            row++;

            if (rows.length === 0) {
                const c1 = sheet.getCell(row, startCol);
                c1.value = 'No roles';
                c1.alignment = { horizontal: 'left', vertical: 'middle' };
                const c2 = sheet.getCell(row, startCol + 1);
                c2.value = 0;
                c2.alignment = { horizontal: 'right', vertical: 'middle' };
                row++;
            } else {
                for (const r of rows) {
                    const c1 = sheet.getCell(row, startCol);
                    c1.value = r.role;
                    c1.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
                    const c2 = sheet.getCell(row, startCol + 1);
                    c2.value = r.count;
                    c2.alignment = { horizontal: 'right', vertical: 'top' };
                    row++;
                }
            }

            row++;
        };

        for (const b of brands) {
            const entry = counts.get(b.id) ?? { brandName: b.name, rows: [] };
            writeBrandBlock(entry.brandName, entry.rows);
        }

        const endRow = row - 1;
        for (let r = 2; r <= endRow; r++) {
            for (let c = startCol; c <= startCol + 1; c++) {
                const cell = sheet.getCell(r, c);
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
            }
        }

        return endRow;
    };

    const endRowVi = writeSide(offsetVi, brandsVi, viCounts);
    writeSide(offsetVde, brandsVde, vdeCounts);

    const uniqueUserIdsVi = new Set<number>();
    for (const ba of brandAgentsVi) {
        uniqueUserIdsVi.add(ba.user_id);
    }

    const userIdToRoleIdAgg = new Map<number, number | undefined>(
        usersVi.map((u) => [u.id, u.custom_role_id]),
    );
    const roleIdToNameAgg = new Map<number, string>(rolesVi.map((r) => [r.id, r.name]));

    const viAggregateMap = new Map<string, number>();
    for (const userId of uniqueUserIdsVi) {
        const roleId = userIdToRoleIdAgg.get(userId);
        if (!roleId) continue;
        const roleName = roleIdToNameAgg.get(roleId);
        if (!roleName) continue;
        viAggregateMap.set(roleName, (viAggregateMap.get(roleName) ?? 0) + 1);
    }

    const viAggregatedRows: BrandRoleRows = Array.from(viAggregateMap.entries())
        .map(([role, count]) => ({ role, count }))
        .sort((a, b) => b.count - a.count || a.role.localeCompare(b.role));

    const writeRoleCountBlock = (
        startRow: number,
        startCol: number,
        title: string,
        rows: BrandRoleRows,
    ): number => {
        let row = startRow;

        sheet.mergeCells(row, startCol, row, startCol + 1);
        const titleCell = sheet.getCell(row, startCol);
        titleCell.value = title;
        titleCell.font = { bold: true };
        titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
        titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EEF2' } };
        row++;

        const roleHdr = sheet.getCell(row, startCol);
        roleHdr.value = 'Role';
        roleHdr.font = { bold: true };
        roleHdr.alignment = { horizontal: 'left', vertical: 'middle' };

        const cntHdr = sheet.getCell(row, startCol + 1);
        cntHdr.value = 'Team members';
        cntHdr.font = { bold: true };
        cntHdr.alignment = { horizontal: 'right', vertical: 'middle' };
        row++;

        if (rows.length === 0) {
            const c1 = sheet.getCell(row, startCol);
            c1.value = 'No roles';
            c1.alignment = { horizontal: 'left', vertical: 'middle' };
            const c2 = sheet.getCell(row, startCol + 1);
            c2.value = 0;
            c2.alignment = { horizontal: 'right', vertical: 'middle' };
            row++;
        } else {
            for (const r of rows) {
                const c1 = sheet.getCell(row, startCol);
                c1.value = r.role;
                c1.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
                const c2 = sheet.getCell(row, startCol + 1);
                c2.value = r.count;
                c2.alignment = { horizontal: 'right', vertical: 'top' };
                row++;
            }
        }

        const endRow = row - 1;
        for (let rr = startRow; rr <= endRow; rr++) {
            for (let cc = startCol; cc <= startCol + 1; cc++) {
                const cell = sheet.getCell(rr, cc);
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
            }
        }

        return endRow;
    };

    writeRoleCountBlock(2, offsetAgg, 'All roles across all brands (VI)', viAggregatedRows);
}

export function generateCommonUsersWorksheet(
    sheet: Worksheet,
    usersVi: UserWithCustomAgentRole[],
    usersVde: UserWithCustomAgentRole[],
    usedSeats: number,
): void {
    sheet.getRow(2).height = 18;

    const columns: string[] = [
        'Name',
        'Mail',
        'Role',
        'Created at',
        'Details',
        'Notes',
        'Tags',
        'Last login in days',
    ];

    sheet.mergeCells(1, 1, 1, columns.length);
    const topHeader = sheet.getCell(1, 1);
    topHeader.value = 'Licenses in both VI and VDE';
    topHeader.alignment = { horizontal: 'center', vertical: 'middle' };
    topHeader.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    topHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF03363D' } };

    const headerPadding = 3;
    columns.forEach((col, i) => {
        const column = sheet.getColumn(1 + i);
        column.width = Math.max(col.length + headerPadding, 12);
    });

    const vdeEmails = new Set(
        usersVde.map((u) => (u.email ?? '').trim().toLowerCase()).filter((e) => e.length > 0),
    );

    const commonUsersViSide = usersVi.filter((u) => {
        const email = (u.email ?? '').trim().toLowerCase();
        return email.length > 0 && vdeEmails.has(email);
    });

    const rows = commonUsersViSide.map((user) => [
        user.name,
        user.email ?? '',
        user.custom_agent_role_name,
        formatDate(user.created_at),
        user.details,
        user.notes,
        (user.tags ?? []).join(', '),
        daysSince(user.last_login_at),
    ]);

    const safeSheetName = sheet.name.replace(/[^A-Za-z0-9_]/g, '_');
    const tableName = `TableCommon_${safeSheetName}`;

    const writeHeaderWithoutTable = (startCol: number) => {
        columns.forEach((name, i) => {
            const cell = sheet.getCell(2, startCol + i);
            cell.value = name;
            cell.font = { bold: true };
            cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        });
    };

    if (rows.length > 0) {
        sheet.addTable({
            name: tableName,
            ref: sheet.getCell(2, 1).address,
            style: { theme: 'TableStyleLight1' },
            columns: columns.map((name) => ({ name, filterButton: true })),
            rows,
        });
    } else {
        writeHeaderWithoutTable(1);
    }

    const applyWrapAndBorder = (
        startRow: number,
        startCol: number,
        rowCount: number,
        colCount: number,
    ) => {
        const endRow = startRow + rowCount - 1;
        const endCol = startCol + colCount - 1;
        for (let r = startRow; r <= endRow; r++) {
            for (let c = startCol; c <= endCol; c++) {
                const cell = sheet.getCell(r, c);
                const currentAlignment = cell.alignment ?? {};
                const isNumericCol = c === endCol;
                const desiredHorizontal: 'left' | 'right' = isNumericCol ? 'right' : 'left';
                cell.alignment = {
                    ...currentAlignment,
                    wrapText: true,
                    horizontal: desiredHorizontal,
                    vertical: 'top',
                };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' },
                };
            }
        }
    };

    const dataStartRow = 3;
    if (rows.length > 0) {
        applyWrapAndBorder(dataStartRow, 1, rows.length, columns.length);
    }

    const labelRow = rows.length > 0 ? 2 + rows.length + 1 : 3;
    const labelDoubleUsedSeats = sheet.getCell(labelRow, 1);
    labelDoubleUsedSeats.value = 'Double used seats';
    labelDoubleUsedSeats.font = { bold: true };

    const valueDoubleUsedSeats = sheet.getCell(labelRow, 2);
    valueDoubleUsedSeats.value = rows.length > 7 ? rows.length - 7 : 0;
    valueDoubleUsedSeats.alignment = { horizontal: 'left', vertical: 'middle' };

    const labelUsedSeats = sheet.getCell(labelRow + 1, 1);
    labelUsedSeats.value = 'Used seats';
    labelUsedSeats.font = { bold: true };

    const valueUsedSeats = sheet.getCell(labelRow + 1, 2);
    valueUsedSeats.value = usedSeats;
    valueUsedSeats.alignment = { horizontal: 'left', vertical: 'middle' };
}
