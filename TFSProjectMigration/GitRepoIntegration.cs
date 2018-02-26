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
        private readonly TfsTeamProjectCollection _tfs;
        private readonly GitRepositoryService _git;
        internal class CommitInfo
        {
            public Guid RepoId { get; set; }
            public string CommitId { get; set; }
            public string Comment { get; set; }
        }

        public GitRepoIntegration(TfsTeamProjectCollection tfs)
        {
            _tfs = tfs;
            _git = new GitRepositoryService();
            _git.Initialize(tfs);
        }

        public IEnumerable<CommitInfo> GetCommits()
        {
            foreach (var repo in _git.QueryRepositories(string.Empty))
            {
                var gitClient = _tfs.GetClient<GitHttpClient>();

                var repoId = repo.Id;

                var criteria = new GitQueryCommitsCriteria();

                int skip = 0;
                bool more;

                do
                {
                    var commitRefs = gitClient.GetCommitsAsync(repoId, criteria, skip).Result;
                    skip += commitRefs.Count;

                    foreach (var commitRef in commitRefs)
                    {
                        var comment = commitRef.CommentTruncated ? gitClient.GetCommitAsync(commitRef.CommitId, repoId).Result.Comment : commitRef.Comment;

                        yield return
                            new CommitInfo
                            {
                                RepoId = repoId,
                                CommitId = commitRef.CommitId,
                                Comment = comment
                            };
                    }

                    more = commitRefs.Count > 0;
                } while (more);

            }

        }

    }
}
