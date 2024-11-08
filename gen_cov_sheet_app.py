import streamlit as st
import pandas as pd
import pg8000  # Import pg8000 instead of psycopg2
import os

# Function to generate coversheets
def generate_coversheets(student_list=[], file_location=""):
    db_connection = pg8000.connect(
        database="postgres",
        user="postgres.yetmtzyyztirghaxnccp",
        password="Keyblade19731998",
        host="aws-0-ap-southeast-1.pooler.supabase.com",
        port=6543
    )

    db_cursor = db_connection.cursor()
    student_list_string = ', '.join(map(str, student_list))

    db_query = f"""SELECT student_list.name,                 
                    student_list.iatc_id, 
                    student_list.class,
                    exam_results.exam,
                    exam_results.score,
                    exam_results.result,
                    exam_results.date
                    FROM exam_results 
                    JOIN student_list ON exam_results.nat_id = student_list.nat_id
                    WHERE student_list.iatc_id IN ({student_list_string}) AND exam_results.score_index = 1
                    ORDER BY exam_results.date ASC
                """
    db_cursor.execute(db_query)
    output_data = db_cursor.fetchall()
    db_cursor.close()
    db_connection.close()

    col_names = ['name', 'iatc_id', 'class', 'exam', 'score', 'result', 'date']
    df = pd.DataFrame(output_data, columns=col_names)

    if not os.path.exists(file_location):
        os.makedirs(file_location)

    for student_id in student_list:
        filtered_df = df[df['iatc_id'] == student_id]
        filtered_df.to_excel(f"{file_location}/{student_id}.xlsx")

# Streamlit interface
st.title("Generate Coversheets")
st.write("Enter a list of student IDs and download the Excel coversheets.")

student_ids_input = st.text_area("Enter Student IDs separated by commas (e.g., 151596, 156756, 154960):")
file_location = st.text_input("Enter the output file path or leave blank to use default location:", value="output_folder")

if st.button("Generate Coversheets"):
    try:
        student_list = [int(id.strip()) for id in student_ids_input.split(",")]
        st.write("Generating coversheets...")

        generate_coversheets(student_list, file_location)

        st.success("Coversheets generated successfully!")
        for student_id in student_list:
            file_path = os.path.join(file_location, f"{student_id}.xlsx")
            with open(file_path, "rb") as file:
                st.download_button(
                    label=f"Download {student_id}.xlsx",
                    data=file,
                    file_name=f"{student_id}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    except Exception as e:
        st.error(f"An error occurred: {e}")
