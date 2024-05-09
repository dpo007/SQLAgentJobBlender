DECLARE @NewJobName NVARCHAR(128) = 'Access Grantor'; -- Name for new job
DECLARE @SourceJobNameSpec NVARCHAR(128) = 'Grant%'; -- Name spec for jobs to combine
DECLARE @ScheduleName NVARCHAR(128) = '7:15AM Daily'; -- Give a name to the schedule

-- Check if the job exists and delete it if it does
IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = @NewJobName)
BEGIN
    EXEC msdb.dbo.sp_delete_job @job_name = @NewJobName;
END

-- Create a temporary table to store the steps from "Grant" jobs
CREATE TABLE #GrantSteps (
    ID INT IDENTITY(1,1),
    JobName NVARCHAR(128),
    StepID INT,
    StepName NVARCHAR(128),
    StepOrder INT,
    Command NVARCHAR(MAX)
)

-- Insert steps from jobs starting with "Grant" into the temporary table
INSERT INTO #GrantSteps (JobName, StepID, StepName, StepOrder, Command)
SELECT
    J.name AS JobName,
    S.step_id AS StepID,
    S.step_name AS StepName,
    ROW_NUMBER() OVER (PARTITION BY J.name ORDER BY S.step_id) AS StepOrder,
    S.command AS Command
FROM msdb.dbo.sysjobs J
JOIN msdb.dbo.sysjobsteps S ON J.job_id = S.job_id
WHERE J.name LIKE @SourceJobNameSpec;

-- Create the new job
DECLARE @NewJobID UNIQUEIDENTIFIER;
EXEC msdb.dbo.sp_add_job @job_name = @NewJobName, @enabled = 1, @job_id = @NewJobID OUTPUT;

-- Add steps to the job from the temporary table
DECLARE @JobName NVARCHAR(128);
DECLARE @StepID INT;
DECLARE @StepName NVARCHAR(128);
DECLARE @StepOrder INT;
DECLARE @Command NVARCHAR(MAX);

DECLARE curSteps CURSOR FOR
SELECT ID, JobName, StepID, StepName, StepOrder, Command
FROM #GrantSteps
ORDER BY JobName, StepOrder;

OPEN curSteps;
FETCH NEXT FROM curSteps INTO @StepID, @JobName, @StepID, @StepName, @StepOrder, @Command;

WHILE @@FETCH_STATUS = 0
BEGIN
    EXEC msdb.dbo.sp_add_jobstep
        @job_id = @NewJobID,
        @step_name = @StepName,
        @step_id = @StepID,
        @subsystem = N'TSQL',
        @command = @Command,
        @on_success_action = 3, -- Goto Next Step
        @on_fail_action = 2; -- Quit with failure

    FETCH NEXT FROM curSteps INTO @StepID, @JobName, @StepID, @StepName, @StepOrder, @Command;
END

CLOSE curSteps;
DEALLOCATE curSteps;

-- Update the last step's "@on_success_action" to 1 (Quit with success)
DECLARE @LastStepID INT;
SELECT TOP 1 @LastStepID = ID
FROM #GrantSteps
ORDER BY ID DESC;

EXEC msdb.dbo.sp_update_jobstep
    @job_id = @NewJobID,
    @step_id = @LastStepID,
    @on_success_action = 1; -- Quit with success

-- Add a schedule to the job to run at 7:15 AM every day
DECLARE @ScheduleID INT;
EXEC msdb.dbo.sp_add_jobschedule
    @job_id = @NewJobID,
    @name = @ScheduleName,
    @freq_type = 4, -- Daily
    @freq_interval = 1, -- Every day
    @active_start_time = 71500; -- 7:15 AM in HHMMSS format

-- Set job to target local server
EXEC msdb.dbo.sp_add_jobserver
    @job_name = @NewJobName,
    @server_name = N'(LOCAL)';

-- Clean up: Drop temporary table
DROP TABLE #GrantSteps;