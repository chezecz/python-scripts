from utilities.get_dates import get_period_dates

def set_task_params():
    first_date, second_date = get_period_dates()
    params = f",,0000,9999,,,{first_date} 00:00:00.000,{second_date} 23:59:59.000,000C~LIT~(0)"
    report = "LM_RPT_JobCosts_Summary_0203"
    report_executable = "LM_RPT_JobCosts_Summary_minmax_0203"
    task_type = "RPT"
    task_params = [report, task_type, report_executable, params]
    return "'" + "','".join(task_params) + "'"
	
def task_query():
	return f"SET NOCOUNT ON; insert into ActiveBGTasks (TaskName, TaskTypeCode, TaskExecutable, TaskParms1) values ({set_task_params()}) select 1"