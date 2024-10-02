require("dotenv").config();
const axios = require("axios");
const ExcelJS = require("exceljs");
const { startOfMonth, endOfMonth, format } = require("date-fns");

const jiraAccessToken = process.env.JIRA_ACCESS_TOKEN;
const assigneeNames = process.env.JIRA_TEAM_MEMBERS.split(",");
const jiraUrl = process.env.JIRA_URL;
const jiraProjectName = process.env.JIRA_PROJECT_NAME;
const jiraStatus = process.env.JIRA_STATUS;
const jiraStatusCategory = process.env.JIRA_STATUS_CATEGORY;
const jiraOrderBy = process.env.JIRA_ORDER_BY;
const jiraOrderDirection = process.env.JIRA_ORDER_DIRECTION;
const jiraTeamName = process.env.JIRA_TEAM_NAME;

const currentDate = new Date();
console.log("currentDate", format(currentDate, "yyyy-MM-dd"));
const monthStart = startOfMonth(currentDate);
const monthEnd = endOfMonth(currentDate);

const workbook = new ExcelJS.Workbook();

const fetchIssuesForAssignee = async (assignee) => {
  const config = {
    method: "get",
    maxBodyLength: Infinity,
    url: `${jiraUrl}/rest/api/2/search?jql=project in (${jiraProjectName}) AND assignee was in (${assignee}) during ("${format(
      monthStart,
      "yyyy-MM-dd"
    )}", "${format(
      monthEnd,
      "yyyy-MM-dd"
    )}") AND status was not in (${jiraStatus}) before "${format(
      monthStart,
      "yyyy-MM-dd"
    )}" AND statusCategory = ${jiraStatusCategory} ORDER BY ${jiraOrderBy} ${jiraOrderDirection}`,
    headers: {
      Authorization: `Bearer ${jiraAccessToken}`,
    },
  };

  try {
    const response = await axios.request(config);
    return response.data.issues;
  } catch (error) {
    console.error(error.message);
    return [];
  }
};

const applyConditionalFill = (value) => {
  if (/done/i.test(value)) {
    return { type: "pattern", pattern: "solid", fgColor: { argb: "ffccffce" } }; // Light green color
  } else if (/won't do/i.test(value) || /closed/i.test(value)) {
    return { type: "pattern", pattern: "solid", fgColor: { argb: "FFD3D3D3" } }; // Light gray color
  }
  return null;
};

const applyHeaderStyles = (worksheet, headers) => {
  const borderStyle = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };

  const customStyles = {
    month: { fgColor: { argb: "ffffff00" } }, // Yellow
    year: { fgColor: { argb: "fff8e5d5" } }, // Light pink
    number_of_task: { fgColor: { argb: "ffdbedf4" } }, // Light blue
  };

  const headerRow = worksheet.getRow(1);

  headers.forEach((header, index) => {
    const cell = headerRow.getCell(index + 1);
    cell.font = { bold: true };
    cell.border = borderStyle;
    if (customStyles[header.key]) {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: customStyles[header.key].fgColor,
      };
    }
  });
  headerRow.height = 30;
  headerRow.alignment = { vertical: "middle" };
  headerRow.commit();
};

const createSharedWorksheet = async (issuesByAssignee) => {
  const worksheet = workbook.addWorksheet("All Issues");

  const headers = [
    { header: "Project", key: "project", width: 20 },
    { header: "Assignee", key: "assignee", width: 20 },
    { header: "Issue Key", key: "key", width: 20 },
    { header: "Issue ID", key: "issue_id", width: 20 },
    { header: "Parent ID", key: "parent_id", width: 20 },
    { header: "Summary", key: "summary", width: 32 },
    { header: "Status", key: "status", width: 15 },
    { header: "Created Date", key: "created", width: 20 },
    { header: "Updated Date", key: "updated", width: 20 },
    { header: "Month-GTV", key: "month", width: 20 },
    { header: "Year-GTV", key: "year", width: 20 },
    { header: "Number of Task", key: "number_of_task", width: 20 },
    { header: "Issue Type", key: "issuetype", width: 20 },
    { header: "Resolution", key: "resolution", width: 20 },
  ];

  worksheet.columns = headers;

  applyHeaderStyles(worksheet, headers);

  worksheet.autoFilter = { from: "A1", to: "N1" };

  issuesByAssignee.forEach((assigneeIssues, index) => {
    assigneeIssues.issues.forEach((issue) => {
      const row = worksheet.addRow({
        project: issue.fields.project.key,
        assignee: assigneeIssues.assignee,
        key: issue.key,
        issue_id: issue.id,
        parent_id: issue.fields.parent ? issue.fields.parent.id : "",
        summary: issue.fields.summary,
        status: issue.fields.status.name,
        created: format(new Date(issue.fields.created), "yyyy-MM-dd HH:mm"),
        updated: format(new Date(issue.fields.updated), "yyyy-MM-dd HH:mm"),
        month: format(new Date(issue.fields.resolutiondate), "MMMM"),
        year: format(new Date(issue.fields.resolutiondate), "yyyy"),
        number_of_task: issue.fields.subtasks?.length || 1,
        issuetype: issue.fields.issuetype.name,
        resolution: issue.fields.resolution?.name || "Unresolved",
      });

      row.eachCell((cell) => {
        cell.border = {
          top: { style: "thin" },
          left: { style: "thin" },
          bottom: { style: "thin" },
          right: { style: "thin" },
        };
      });

      // Apply conditional fill to the "Status" cell
      const statusCell = row.getCell("status");
      const statusFillStyle = applyConditionalFill(issue.fields.status.name);
      if (statusFillStyle) {
        statusCell.fill = statusFillStyle;
      }

      // Apply conditional fill to the "Resolution" cell
      const resolutionCell = row.getCell("resolution");
      const fillStyle = applyConditionalFill(
        issue.fields.resolution?.name || ""
      );
      if (fillStyle) {
        resolutionCell.fill = fillStyle;
      }
    });

    if (index < issuesByAssignee.length - 1) {
      const emptyRow = worksheet.addRow({
        project: "",
        assignee: "",
        key: "",
        issue_id: "",
        parent_id: "",
        summary: "",
        status: "",
        created: "",
        updated: "",
        month: "",
        year: "",
        issuetype: "",
        resolution: "",
      });
      emptyRow.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ffff00" },
      };

      emptyRow.commit();
    }
  });

  worksheet.font = { name: "Calibri", size: 11 };
};

const processAssignees = async () => {
  const issuesByAssignee = await Promise.all(
    assigneeNames.map(async (assignee) => {
      const issues = await fetchIssuesForAssignee(assignee);
      return { assignee, issues };
    })
  );

  await createSharedWorksheet(issuesByAssignee);

  await workbook.xlsx.writeFile(
    `${jiraTeamName}_Report_${format(currentDate, "yyyy-MM-dd")}.xlsx`
  );
  console.log("Data exported to Excel file successfully.");
};

// main
processAssignees();
