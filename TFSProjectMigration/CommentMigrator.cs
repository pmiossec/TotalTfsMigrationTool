using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace TFSProjectMigration
{
    class CommentMigrator
    {
        private readonly Project _destinationProject;
        private readonly bool _areVersionHistoryCommentsIncluded;
        private readonly bool _shouldWorkItemsBeLinkedToGitCommits;
        private readonly GitRepoIntegration _gitRepoIntegration;
        private Dictionary<string, GitRepoIntegration.CommitInfo> _commitsByChangeset;

        const int MaxHistoryLength = 1048576;

        private enum RevisionMigrateAction
        {
            Skip,
            MigrateComment,
            MigrateLink
        }

        public CommentMigrator(TfsTeamProjectCollection tfs, Project destinationProject, bool areVersionHistoryCommentsIncluded, bool shouldWorkItemsBeLinkedToGitCommits)
        {
            _destinationProject = destinationProject;
            _areVersionHistoryCommentsIncluded = areVersionHistoryCommentsIncluded;
            _shouldWorkItemsBeLinkedToGitCommits = shouldWorkItemsBeLinkedToGitCommits;
            if (tfs != null)
                _gitRepoIntegration = new GitRepoIntegration(tfs);
        }

        private string getCommitUri(string changeset)
        {
            if (_commitsByChangeset == null)
                initCommitsByChangeset();

            if (_commitsByChangeset.ContainsKey(changeset))
            {
                return CreateCommitUriFromCommit(_commitsByChangeset[changeset]);
            }
            return null;

            string CreateCommitUriFromCommit(GitRepoIntegration.CommitInfo commit)
            {
                return $"vstfs:///Git/Commit/{_destinationProject.Guid}%2F{commit.RepoId}%2F{commit.CommitId}";
            }
    }

        private void initCommitsByChangeset()
        {
            _commitsByChangeset = new Dictionary<string, GitRepoIntegration.CommitInfo>();

            foreach (var commitInfo in _gitRepoIntegration.GetCommits())
            {
                var changeset = GetChangesetFromComment(commitInfo.Comment);
                if (!string.IsNullOrWhiteSpace(changeset) && !_commitsByChangeset.ContainsKey(changeset))
                    _commitsByChangeset.Add(changeset, commitInfo);
            }

            string GetChangesetFromComment(string comment)
            {
                const string gitTfsId = "git-tfs-id:";
                var gitTfsIdIndex = comment.IndexOf(gitTfsId, StringComparison.Ordinal);
                if (gitTfsIdIndex < 0)
                    return null;
                var changeset = comment.Substring(gitTfsIdIndex + gitTfsId.Length).Split(';').Last();

                if (string.IsNullOrEmpty(changeset) || changeset.Length < 2)
                    return null;

                changeset = changeset.Substring(1);

                if (new Regex(@"^\d+$").IsMatch(changeset))
                    return changeset;

                return null;
            }
        }

        public void MigrateComments(WorkItem sourceWorkItem, WorkItem targetWorkItem)
        {
            if (!_areVersionHistoryCommentsIncluded && !_shouldWorkItemsBeLinkedToGitCommits)
                return;

            var migratedLinks = new List<string>();

            Debug.WriteLine("Work Item: " + sourceWorkItem.Id);

            foreach (Revision revision in sourceWorkItem.Revisions)
            {
                switch (ShouldRevisionBeMigrated(revision))
                {
                    case RevisionMigrateAction.MigrateComment:
                        var history = ExtractHistoryFromRevision(revision);

                        if (history.Length < MaxHistoryLength)
                        {
                            SaveHistory(targetWorkItem, history, revision);
                        }

                        break;
                    case RevisionMigrateAction.MigrateLink:
                        history = ExtractHistoryFromRevision(revision);
                        foreach (Link x in revision.Links)
                        {
                            if (x is ExternalLink el && !migratedLinks.Contains(el.LinkedArtifactUri))
                            {
                                migratedLinks.Add(el.LinkedArtifactUri);
                                var changeset = el.LinkedArtifactUri.Split('/').Last();
                                var commitUri = getCommitUri(changeset);
                                if (!string.IsNullOrWhiteSpace(commitUri))
                                {
                                    targetWorkItem.Links.Add(new ExternalLink(_destinationProject.Store.RegisteredLinkTypes[3], commitUri));
                                }
                            }

                        }

                        SaveHistory(targetWorkItem, history, revision);

                        break;
                }
            }

            //  
            void SaveHistory(WorkItem workItem, string history, Revision revision)
            {
                workItem.Fields["Changed By"].Value = revision.Fields["Changed By"].Value;
                workItem.History = history;
                workItem.Save();
            }

            string ExtractHistoryFromRevision(Revision revision)
            {
                var history =
                    $"{revision.Fields["History"].Value}<br><p><FONT color=#808080 size=1><EM>Changed date: {revision.Fields["Changed Date"].Value}</EM></FONT></p>";
                RemoveImagesIfHistoryIsTooLong(ref history);
                return history;
            }


        }

        RevisionMigrateAction ShouldRevisionBeMigrated(Revision revision)
        {
            if (revision.Fields["Changed By"].Value?.ToString() == "_SYSTFSBuild")
                return RevisionMigrateAction.Skip;

            var history = revision.Fields["History"].Value?.ToString();

            if (string.IsNullOrWhiteSpace(history))
                return RevisionMigrateAction.Skip;

            if (history.StartsWith("Associated with changeset"))
            {
                if (_shouldWorkItemsBeLinkedToGitCommits)
                    return RevisionMigrateAction.MigrateLink;
                return RevisionMigrateAction.Skip;
            }

            if (_areVersionHistoryCommentsIncluded)
                return RevisionMigrateAction.MigrateComment;
            return RevisionMigrateAction.Skip;
        }

        void RemoveImagesIfHistoryIsTooLong(ref string history)
        {
            int ImageStart(string s)
            {
                return s.IndexOf("<img", StringComparison.Ordinal);
            }

            int ImageEnd(string s, int imageStart)
            {
                return s.IndexOf(">", imageStart, StringComparison.Ordinal);
            }

            {
                var imageStart = ImageStart(history);

                while (history.Length >= MaxHistoryLength && imageStart >= 0)
                {
                    var imageEnd = ImageEnd(history, imageStart);
                    if (imageEnd < 0)
                        return;
                    history = history.Remove(imageStart, imageEnd - imageStart);
                    history = history.Insert(imageStart, "<p><FONT color=#808080 size=1><EM>(image removed)</EM></FONT></p>");
                    imageStart = ImageStart(history);
                }
            }
        }

    }
}
