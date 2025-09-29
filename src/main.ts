import { Workbook } from 'exceljs';

import {
    authVde,
    authVi,
    baseUrlVde,
    baseUrlVi,
    fetchAllBrandAgents,
    fetchAllBrands,
    fetchAllCustomAgentRoles,
    fetchAllUsers,
} from './api';
import { filterVorwerkInternational } from './brand';
import {
    generateBrandRoleCountWorksheet,
    generateCommonUsersWorksheet,
    generateUserWorksheet,
} from './excel';
import {
    countUsedSeats,
    filterAgentsInactive45Days,
    filterBrandRoleCount,
    filterCommonUsers,
    filterFunctionalUsers,
    filterLightAgentActiveButNotLightAgent,
    filterSuspendedAgentsNotLightAgents,
    mapUsersWithCustomAgentRoleName,
} from './user';

async function fetchAndSave(): Promise<void> {
    try {
        const [
            usersVi,
            usersVde,
            customAgentRolesVi,
            customAgentRolesVde,
            brandsVi,
            brandsVde,
            brandAgentsVi,
            brandAgentsVde,
        ] = await Promise.all([
            fetchAllUsers(baseUrlVi, authVi),
            fetchAllUsers(baseUrlVde, authVde),
            fetchAllCustomAgentRoles(baseUrlVi, authVi),
            fetchAllCustomAgentRoles(baseUrlVde, authVde),
            fetchAllBrands(baseUrlVi, authVi),
            fetchAllBrands(baseUrlVde, authVde),
            fetchAllBrandAgents(baseUrlVi, authVi),
            fetchAllBrandAgents(baseUrlVde, authVde),
        ]);

        const workbook = new Workbook();

        generateUserWorksheet(
            workbook.addWorksheet('Inactive 45 Days or more'),
            mapUsersWithCustomAgentRoleName(
                filterAgentsInactive45Days(usersVi),
                customAgentRolesVi,
            ),
            mapUsersWithCustomAgentRoleName(
                filterAgentsInactive45Days(usersVde),
                customAgentRolesVde,
            ),
        );
        generateUserWorksheet(
            workbook.addWorksheet('Suspended agents not LA'),
            mapUsersWithCustomAgentRoleName(
                filterSuspendedAgentsNotLightAgents(usersVi),
                customAgentRolesVi,
            ),
            mapUsersWithCustomAgentRoleName(
                filterSuspendedAgentsNotLightAgents(usersVde),
                customAgentRolesVde,
            ),
        );
        generateUserWorksheet(
            workbook.addWorksheet('Light agent active not LA'),
            mapUsersWithCustomAgentRoleName(
                filterLightAgentActiveButNotLightAgent(usersVi),
                customAgentRolesVi,
            ),
            mapUsersWithCustomAgentRoleName(
                filterLightAgentActiveButNotLightAgent(usersVde),
                customAgentRolesVde,
            ),
        );
        generateUserWorksheet(
            workbook.addWorksheet('Functional users'),
            mapUsersWithCustomAgentRoleName(filterFunctionalUsers(usersVi), customAgentRolesVi),
            mapUsersWithCustomAgentRoleName(filterFunctionalUsers(usersVde), customAgentRolesVde),
        );
        generateBrandRoleCountWorksheet(
            workbook.addWorksheet('Brand role count'),
            brandsVi,
            filterVorwerkInternational(brandsVde),
            brandAgentsVi,
            brandAgentsVde,
            filterBrandRoleCount(usersVi),
            filterBrandRoleCount(usersVde),
            customAgentRolesVi,
            customAgentRolesVde,
        );
        generateCommonUsersWorksheet(
            workbook.addWorksheet('Agents in both instances'),
            mapUsersWithCustomAgentRoleName(filterCommonUsers(usersVi), customAgentRolesVi),
            mapUsersWithCustomAgentRoleName(filterCommonUsers(usersVde), customAgentRolesVde),
            countUsedSeats(usersVi, usersVde),
        );

        await workbook.xlsx.writeFile('user_report.xlsx');
        console.log('File saved!');
    } catch (err) {
        console.error(err);
    }
}

void fetchAndSave();
