from airflow.sdk import dag
from airflow.providers.standard.operators.trigger_dagrun import TriggerDagRunOperator


@dag
def dag_orchestrate_parent():
    '''
    @task
    def first_orchestrator_task():
        first_orchestrator_dag()
    
    @task
    def second_orchestrator_task():
        second_orchestrator_dag()
    
    first_orchestrator_task() >> second_orchestrator_task()
    '''

    trigger_first_dag = TriggerDagRunOperator(
        task_id="trigger_first_orchestrator_dag", 
        trigger_dag_id="first_orchestrator_dag",
        wait_for_completion=True,
    )

    trigger_second_dag = TriggerDagRunOperator(
        task_id="trigger_second_orchestrator_dag",
        trigger_dag_id="second_orchestrator_dag",
        wait_for_completion=True,
    )

    trigger_first_dag >> trigger_second_dag

dag_orchestrate_parent()
