# Total TFS Migration Tool

Improved clone over https://totaltfsmigration.codeplex.com used to migrate workitems from one TFS project to another.

> Note: Description below is the [original description from codeplex](https://totaltfsmigration.codeplex.com/).

## Project Description

The Total TFS Migration Tool is a tool developed by Geveo Australasia to facilitate moving data between two TFS Projects. These project can reside in one TFS Server instance or in two TFS Server instances. Currently tool support migration of Work items, Shared queries, Iterations, Areas Test Plans, Test Suits and Work flows. All the Migration Process is done through the APIs of TFS. This is intended as a generic tool which can be used by others, but users can customize it to match their requirements.

## Vision

The TFS Migration tool is for any person or company which need a TFS project migration tool. Main goal of the project is to write a simple application that can replicate same development environment of a TFS Project in another TFS Project. This tool was developed as an internal tool to transfer work items between projects in different TFS Servers which runs different TFS versions. As there wasn’t a generic tool to done a similar work, we thought developing a generic tool might help others who have the same need. Building a total generic tool for TFS Migration is almost impossible as every company that use TFS as a source control software and as a team management software, have customized it to suit their needs and development methodologies. But as the source code is available anyone can customize this tool to match their requirements.

Feedback from end users will also be critical to make this more generic, so please help us by sharing your experiences and feedback.

## Scope

At this point in time, the TFS Migration Tool is scoped to only support the following TFS features:

* Work Item Migration - This includes the migration of all work items, fields, in-use areas, iterations, links, attachments, shared queries.
* Test Plan Migration – This includes migration of all test plans and test suits.

## Features

* Copy custom fields of relevant work item types in source project to target project.
* Copy workflows from source TFS project to target TFS project.
* Map the mismatched source fields with target fields.
* Logging mechanism to track entire migration process.

## Preparing For Migration

Before starting the Migration Tool you need to customize the Work Item Definition so you will not have problems during the Migration.

* Remove or Modify <REQUIRED> tags to make sure you will not have validation errors.
* Make sure the Users names are equal to the Users Name in TFS(Add all users to new project would be smart choice).  
  Or you can remove the <VALIDUSER> from AssignTo field definition if you don't want to compare User Names.
* Make sure that fields or global list have the values your about to add. You can replace <ALLOWEDVALUES> with <SUGGESTEDVALUES> to allow any value.

## Limitations

* No Version Control – Project source code will not be transferred.
* Work item history will not be transferred.

