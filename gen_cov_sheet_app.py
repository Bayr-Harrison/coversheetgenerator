import streamlit as st
import pandas as pd
import pg8000
from io import BytesIO
import zipfile

# Function to generate coversheets and save them to a zip file
def generate_coversheets_zip(student_list=[]):
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

    # Create an in-memory ZIP file
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for student_id in student_list:
            filtered_df = df[df['iatc_id'] == student_id]

            # Save each filtered dataframe to an in-memory buffer
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name=str(student_id))
            
            # Save Excel file to the zip
            excel_filename = f"{student_id}.xlsx"
            excel_buffer.seek(0)
            zip_file.writestr(excel_filename, excel_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer

# Streamlit interface
st.title("Generate Coversheets")
st.write("Enter a list of student IDs and download the Excel coversheets.")

student_ids_input = st.text_area("Enter Student IDs separated by commas (e.g., 151596, 156756, 154960):")

if st.button("Generate Coversheets"):
    try:
        student_list = [int(id.strip()) for id in student_ids_input.split(",")]
        st.write("Generating coversheets...")

        # Generate the zip file in memory
        zip_file = generate_coversheets_zip(student_list)

        # Offer the zip file for download
        st.download_button(
            label="Download All Coversheets as ZIP",
            data=zip_file,
            file_name="coversheets.zip",
            mime="application/zip"
        )
        st.success("Coversheets zip generated successfully!")

    except Exception as e:
        st.error(f"An error occurred: {e}")
