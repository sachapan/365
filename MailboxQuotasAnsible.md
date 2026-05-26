# This is a concept based on chapgpt generated content.  completely untested at the moment.

You are an ansible and ms365 expert.  Create an ansible playbook that runs on a linux server to query ms365 and report all mailboxes near their quota limit.
`
---
- name: Report MS365 mailboxes near quota limits
  hosts: localhost
  gather_facts: false

  vars:
    tenant_id: "YOUR_TENANT_ID"
    client_id: "YOUR_APP_CLIENT_ID"
    client_secret: "YOUR_APP_CLIENT_SECRET"

    warning_threshold_percent: 85

    graph_scope: "https://graph.microsoft.com/.default"

  tasks:

    - name: Obtain OAuth token from Microsoft identity platform
      ansible.builtin.uri:
        url: "https://login.microsoftonline.com/{{ tenant_id }}/oauth2/v2.0/token"
        method: POST
        body_format: form-urlencoded
        body:
          client_id: "{{ client_id }}"
          client_secret: "{{ client_secret }}"
          scope: "{{ graph_scope }}"
          grant_type: client_credentials
        return_content: true
      register: auth_response

    - name: Store access token
      ansible.builtin.set_fact:
        access_token: "{{ auth_response.json.access_token }}"

    - name: Query all licensed users
      ansible.builtin.uri:
        url: "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,assignedLicenses&$top=999"
        method: GET
        headers:
          Authorization: "Bearer {{ access_token }}"
          Content-Type: "application/json"
        return_content: true
      register: users_response

    - name: Build mailbox candidate list
      ansible.builtin.set_fact:
        mailbox_users: >-
          {{
            users_response.json.value
            | selectattr('assignedLicenses', 'defined')
            | selectattr('assignedLicenses', '!=', [])
            | list
          }}

    - name: Query mailbox usage reports
      ansible.builtin.uri:
        url: "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='D7')"
        method: GET
        headers:
          Authorization: "Bearer {{ access_token }}"
        return_content: true
      register: mailbox_report_raw

    - name: Parse CSV mailbox report
      ansible.builtin.set_fact:
        mailbox_report_lines: "{{ mailbox_report_raw.content.splitlines() }}"

    - name: Convert mailbox CSV into structured objects
      ansible.builtin.set_fact:
        mailbox_report: "{{ mailbox_report | default([]) + [parsed_item] }}"
      vars:
        fields: "{{ item.split(',') }}"
        parsed_item:
          user_principal_name: "{{ fields[2] | trim('\"') }}"
          display_name: "{{ fields[1] | trim('\"') }}"
          storage_used_mb: "{{ ((fields[5] | trim('\"') | int) / 1048576) | round(2) }}"
          prohibit_send_receive_quota_mb: "{{ ((fields[8] | trim('\"') | int) / 1048576) | round(2) }}"
      loop: "{{ mailbox_report_lines[1:] }}"
      when:
        - item | length > 0
        - item is search('@')

    - name: Calculate quota utilization
      ansible.builtin.set_fact:
        quota_results: >-
          {{
            quota_results | default([]) +
            [
              {
                'display_name': item.display_name,
                'user_principal_name': item.user_principal_name,
                'used_mb': item.storage_used_mb,
                'quota_mb': item.prohibit_send_receive_quota_mb,
                'percent_used': (
                  (
                    item.storage_used_mb | float /
                    item.prohibit_send_receive_quota_mb | float
                  ) * 100
                ) | round(2)
              }
            ]
          }}
      loop: "{{ mailbox_report }}"
      when:
        - item.prohibit_send_receive_quota_mb | float > 0

    - name: Filter mailboxes near quota
      ansible.builtin.set_fact:
        near_quota_mailboxes: >-
          {{
            quota_results
            | selectattr('percent_used', '>=', warning_threshold_percent)
            | sort(attribute='percent_used', reverse=true)
            | list
          }}

    - name: Display mailboxes near quota
      ansible.builtin.debug:
        msg: |
          Display Name : {{ item.display_name }}
          UPN          : {{ item.user_principal_name }}
          Used (MB)    : {{ item.used_mb }}
          Quota (MB)   : {{ item.quota_mb }}
          Used %       : {{ item.percent_used }}
      loop: "{{ near_quota_mailboxes }}"

    - name: Fail if any mailbox exceeds threshold
      ansible.builtin.fail:
        msg: "One or more mailboxes are above {{ warning_threshold_percent }}% quota usage."
      when: near_quota_mailboxes | length > 0`


  ##  Required Microsoft Graph API permissions
Application permissions required:
Reports.Read.All
User.Read.All
Grant admin consent after assigning permissions.
Azure App Registration
Create an app registration in Microsoft Entra Admin Center
Microsoft Graph Mailbox Usage Report Reference
Microsoft Graph getMailboxUsageDetail API

