from airflow.sdk import dag, task
from pendulum import datetime, duration
from airflow.timetables.trigger import DeltaTriggerTimetable

@dag(
    dag_id="delta_schedule_dag",
    start_date=datetime(2026, 7, 11, tz="Asia/Bangkok"),
    schedule=DeltaTriggerTimetable(duration(days=3)),
    end_date = datetime(2026, 7, 20,tz="Asia/Bangkok" ),
    is_paused_upon_creation=False,
    catchup=True 
)
def delta_schedule_dag():

    @task.python
    def first_task():
        print("This is the first task")

    @task.python
    def second_task():
        print("This is the second task")

    @task.python
    def third_task():
        print("This is the third task")

    first = first_task()
    second = second_task()
    third = third_task()

    first >> second >> third


delta_schedule_dag()