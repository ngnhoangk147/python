from jira import JIRA

jira_connection = JIRA(
    token_auth="{token_information}",
    validate=False,
    server="{httplink}"
)

issue_dict = {
    'project': {'key': 'ITDHITCERT'},
    'summary': 'Testing issue poc',
    'description': 'Detailed ticket description.',
    'issuetype': {'name': 'Task'},
}

new_issue = jira_connection.create_issue(fields=issue_dict)