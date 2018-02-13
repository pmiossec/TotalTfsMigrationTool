using System;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Git.Client;
using Microsoft.TeamFoundation.SourceControl.WebApi;

namespace TFSProjectMigration
{
    class GitRepoIntegration
    {
        public GitRepoIntegration(TfsTeamProjectCollection tfs)
        {
            var git = new GitRepositoryService();
            git.Initialize(tfs);

            foreach (var repo in git.QueryRepositories(string.Empty))
            {
                Debug.WriteLine("Git Repo {0}, id {1}", repo.Name, repo.Id);

                GitHttpClient gitClient = tfs.GetClient<GitHttpClient>();

                Guid repoId = repo.Id; 

                // Get no more than 10 commits
                GitQueryCommitsCriteria criteria = new GitQueryCommitsCriteria()
                {
                    Top = 10
                };

                List<GitCommitRef> commits = gitClient.GetCommitsAsync(repoId, criteria).Result;

                foreach (GitCommitRef commit in commits)
                {
                    Debug.WriteLine("{0} by {1} ({2})", commit.CommitId, commit.Committer.Email, commit.Comment);
                }


            }
        }
    }
}
