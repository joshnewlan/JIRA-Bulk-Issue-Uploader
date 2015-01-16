import xlrd, re
from xlrd import xldate_as_tuple
from jira.client import JIRA

username = ""
password = ""

some_jira = JIRA(options={'server': 'http://url.com'},basic_auth=(username, password))
data = xlrd.open_workbook('/path/to/excelsheet.xlsx')
table = data.sheet_by_name(u'Sheet1')

class Issue(object):
    summary = 0
    project = 1
    description = 2
    issuetype = 3
    components = 4
    affects_version = 5
    fix_versions = 6
    assignee = 7
    estimated_time = 8

    def __init__(self, fields):
        requires = ((self.summary, "summary"), (self.project, "project"), (self.issuetype, "issue_type"), (self.components, "components"))
        for require, require_name in requires:
            if not fields[require]:
                raise Exception("Required field={} is empty or None".format(require_name))
    
        self.summary = fields[self.summary]
        self.description = fields[self.description]
        if self.description == "":
            self.description = self.summary
    
        self.project = fields[self.project]
        self.issuetype = fields[self.issuetype]
        self.components = fields[self.components]
        self.affects_version = fields[self.affects_version]
        self.fix_versions = fields[self.fix_versions]
        self.assignee = fields[self.assignee]
        self.estimated_time = time_to_seconds(fields[self.estimated_time])


def time_to_seconds(str):
    pattern=r'(?:(?P<w>\d+){1}[w|W]{1})?(?:(?P<d>\d+){1}[d|D]{1})?(?:(?P<h>\d+){1}[h|H]{1})?'
    m=re.match(pattern=pattern, string=str)
    timedict = m.groupdict()
    week = 0
    hour = 0
    day = 0
    if timedict["w"]!= None:
        week = timedict["w"]
    else:
        week = 0
    if timedict["h"]!= None:
        hour = timedict["h"]
    else:
        hour = 0
    if timedict["d"]!= None:
        day = timedict["d"]
    else:
        day = 0
    second = int(day)*28800+int(hour)*3600+int(week)*5*28800
    return second


def create_issuelist(excel_table):
    nrows = excel_table.nrows
    issuelist = []
    for r in xrange(1, nrows):
        issuedata = excel_table.row_values(r)
        issue = Issue(issuedata)
        print issue.summary
        issuelist.append(issue)
    return issuelist


def create_issues(issuelist):
    for issue in issuelist:
        query_temp = 'project={} and summary~"{}" and fixVersion="{}" and affectedVersion="{}"'
        stories = some_jira.search_issues(
            query_temp.format(issue.project, issue.summary, issue.fix_versions, issue.affects_version)
        )
        if stories:
            print "Duplicate"
            continue

        issue_dict =  {
            'project': {'key': issue.project},
            'summary': issue.summary,
            'description': issue.description,
            'issuetype': {'name': issue.issuetype},
            'components':[{'name': issue.components}],
            'versions':[{'name': issue.affects_version}],
            'fixVersions':[{'name': issue.fix_versions}],
        }  
        new_issue = some_jira.create_issue(fields=issue_dict)
        story = some_jira.search_issues(
            query_temp.format(issue.project, issue.summary, issue.fix_versions, issue.affects_version)
        )
        if story:
            print story[0]

            if issue.assignee:
                story[0].update(assignee={'name': issue.assignee})
            if issue.estimated_time:
                story[0].update(timetracking={'originalEstimate': issue.estimated_time})


def main():
    all_issues = create_issuelist(table)
    create_issues(all_issues)


if __name__ == "__main__":
    main()  
