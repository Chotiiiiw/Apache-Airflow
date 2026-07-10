from airflow.sdk import dag, task
from airflow.providers.standard.operators.bash import BashOperator

@dag(
  dag_id="xcom_dag_auto"
)
def xcom_dag_auto():

    @task.python
    def first_task():
        print("Extracting data... This is the first task")
        fetched_data = {"data": [1,2,3,4,5]}
        return fetched_data
    
    @task.python
    def second_task(data:dict):
        print("Transforming data... This is the second task")
        fetched_data = data['data']
        transformed_data = fetched_data*2
        transformed_data_dict = {'data': transformed_data}
        return transformed_data_dict

    @task.python
    def third_task(data:dict):
        load_data = data
        return load_data


    #Defining task dependencies
    first = first_task()
    second = second_task(first)
    third = third_task(second)

    ##We can ignore this since airflow is smart. first >> second >> third

xcom_dag_auto()
