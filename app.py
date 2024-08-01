import streamlit as st
import pandas as pd
import os

def drop_duplicates(df):
    old_len = len(df)
    df.drop_duplicates(inplace=True,
                       subset=['Mã NV', 'Họ và tên', 'Chức danh', 
                               'Chi nhánh', 'Phòng ban', 'Ngày tham gia', 
                               'Ngày hoàn thành', 'Điểm bài học', 'Điểm thi',
                               'Ghi chú'])
    num_dups = old_len  - len(df)
    return df, num_dups

def add_course_name(df, course_name):
    df['Tên khoá học'] = course_name
    return df

def concat_df(df_list):
    return pd.concat(df_list)

def click_gen_button():
    st.session_state.clicked = True

if __name__ == '__main__':
    st.title('Tutorial')
    st.write('This tool is used to generate a summary report file based on multiple input reports.')
    st.write('The user need to follow the following steps:')
    st.write('1. Uploading all necessary ".xlsx" report files in the "File Uploading" section.')
    st.write('2. Checking the number of files and duplicates in the "Description" section.')
    st.write('3. Clicking the "Generate Final Report" button to start the process.')
    st.write('4. Waiting and Download the file after finishing.')

    st.title('Files Uploading')
    # st.write('Please upload all your report files here in "xlsx" format')

    uploaded_files = st.file_uploader(
        "Please upload all your report files here in '.xlsx' format", accept_multiple_files=True, type = {'xlsx'}
    )

    count = 0
    df_list = []

    if uploaded_files:
        st.title('Description')

        for uploaded_file in uploaded_files:
            df = pd.DataFrame(pd.read_excel(uploaded_file)) 

            # drop duplicates
            df, num_dups = drop_duplicates(df)

            # add course name
            course_name = uploaded_file.name[:-5]
            df = add_course_name(df, course_name)

            # print number of duplicates if existed
            if num_dups > 0:
                st.write(f'Found {num_dups} duplicates in {course_name}')

            # append df to list
            df_list.append(df)

            # count number of courses
            count += 1
        
        st.write(f"**Total files uploaded: {count}**")
            

        # save name
        save_name = "final_report.xlsx"
        
        st.title('Final Report Generation')

        # if there are files in the list
        if df_list:
            # start generating report
            if st.button('Generate Final Report'):
                st.write('Generating Final Report ...')
                # concat files
                final_df = concat_df(df_list)

                # save df to excel file
                final_df.to_excel(save_name, sheet_name="Final Report", index=False)

                # open and read saved excel file
                with open(save_name, "rb") as template_file:
                    template_byte = template_file.read()

                # remove saved file
                os.remove(save_name)

                st.write('Finish!')
                st.download_button(label='Download Final Report File',
                                    data=template_byte,
                                    file_name= save_name)

