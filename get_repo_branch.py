import requests
def get_repos_and_branch(username, token):
    base_url = 'https://api.github.com'
    headers = {
        'Authorization': f'token {token}',
    }
    repos_url = f'{base_url}/users/{username}/repos'
    response = requests.get(repos_url, headers=headers)
    repos_and_branch = {}
    if response.status_code == 200:
        repos = response.json()
        for repo in repos:
            repo_name = repo['name']
            branches_url = f"{base_url}/repos/{username}/{repo_name}/branches"
            branches_response = requests.get(branches_url, headers=headers)
            if branches_response.status_code == 200:
                branches = branches_response.json()
                branches_name = [branch['name'] for branch in branches]
                repos_and_branch[repo_name] = branches_name
            else:
                print(f'Failed to fetch branches for {repo_name}: {branches_response.status_code}')

    else:
        print(f'Failed : {response.status_code}')
    return repos_and_branch

username = 'liuziang1210'

token = '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'
result = get_repos_and_branch(username, token)
print(result)