import streamlit as st
import pandas as pd
from io import BytesIO
import time
import re
import openpyxl

def process_file(uploaded_file, selected_columns):
    columns_ordered = selected_columns.copy()

    df = pd.read_excel(uploaded_file, sheet_name=0)
    columns_found = {}

    max_row_index = min(10, len(df))
    largest_data_start_row = 0

    # Find the data row for each column selected by the user
    for row_index in range(max_row_index):
        for col_index in range(df.shape[1]):
            cell_value = df.iloc[row_index, col_index]
            if cell_value in selected_columns:
                data_start_row = row_index + 1
                largest_data_start_row = max(largest_data_start_row, data_start_row)
                columns_found[cell_value] = col_index
                selected_columns.remove(cell_value)
                if not selected_columns:
                    break
        if not selected_columns:
            break

    # Create new dataframe with selected columns
    for col_name, col_index in columns_found.items():
        data_series = df.iloc[largest_data_start_row:, col_index].reset_index(drop=True)
        columns_found[col_name] = data_series

    new_df = pd.DataFrame(columns_found)
    
    new_df = new_df.dropna()
    new_df = new_df.reset_index(drop=True)
    new_df = new_df[~new_df['姓名'].apply(lambda x: str(x).isdigit())]
    new_df = new_df[columns_ordered]
    
    return new_df

def find_possible_columns(uploaded_file):
    df = pd.read_excel(uploaded_file, sheet_name=0)
    possible_columns = set()

    chinese_char_pattern = re.compile(r'^[\u4e00-\u9fff]+$')

    max_row_index = min(10, len(df))

    # Iterate through the first ten rows and each column to collect unique values
    for row_index in range(max_row_index):
        for col_index in range(df.shape[1]):
            cell_value = df.iloc[row_index, col_index]
            if pd.notna(cell_value):  # Check if the cell is not NaN
                cell_value_str = str(cell_value)  # Convert to string
                if chinese_char_pattern.match(cell_value_str):  # Only include if it contains only Chinese characters
                    possible_columns.add(cell_value_str)

    return list(possible_columns)

def export_file(df, selected_columns=None, mode=None):

    output = BytesIO()

    if mode == "combine":
        # Adjust the columns based on user selection and bank format
        export_df = df[selected_columns]
        
        # Use a BytesIO object to store the Excel file in memory
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            export_df.to_excel(writer, index=False)
        
        # Set the file position to the beginning before returning
        output.seek(0)
    elif mode == "separate":
        # Provide download option
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

    
    return output

def make_unique(column_names):
    seen = {}
    new_cols = []
    for col in column_names:
        if pd.isna(col):
            col = 'Unnamed'
        count = seen.get(col, 0)
        if count:
            new_col = f"{col}.{count}"
            new_cols.append(new_col)
        else:
            new_cols.append(col)
        seen[col] = count + 1
    return new_cols

def main():
    st.title("博达 - 工作表工具箱")

    tab1, tab2, tab3, tab4 = st.tabs(["多表合一", "单表拆分", "数据稽核", "数据可视化"])

    with tab1:
        uploaded_files = st.file_uploader("选择需要合并的Excel表格", type=["xlsx", "xls"], accept_multiple_files=True)
        
        if uploaded_files:

            possible_columns = find_possible_columns(uploaded_files[0])
            current_default = ['姓名', '身份证号', '应付工资']

            selected_columns = st.multiselect("选择要处理的列", possible_columns, default=current_default)

            process_button = st.button("处理文件")
            
            if process_button:
                total_files = len(uploaded_files)
                progress_bar = st.progress(0)
                status_text = st.empty()
                dfs = []
                total_time = 0
                success_count = 0
                fail_count = 0

                for i, file in enumerate(uploaded_files):
                    try:
                        start_time = time.time()

                        # Process the file with the selected columns
                        df = process_file(file, selected_columns.copy())  # Pass a copy to avoid altering original
                        end_time = time.time()
                        process_time = end_time - start_time
                        total_time += process_time
                        dfs.append(df)
                        success_count += 1
                        st.toast(f"成功处理文件 '{file.name}' \n\n"
                                f"耗时: {process_time:.2f} 秒, 合并行数：{len(df)}", icon="✅")
                    except Exception as e:
                        fail_count += 1
                        st.error(f"处理文件 '{file.name}' 时出错: {str(e)}", icon="🚨")

                    progress = (i + 1) / total_files
                    progress_bar.progress(progress)
                    status_text.text(f"已处理文件: {i + 1}/{total_files} | 成功: {success_count} | 失败: {fail_count}")
                
                if dfs:
                    combined_df = pd.concat(dfs, ignore_index=True)

                    numeric_columns = combined_df.columns
                    selected_numeric_column = st.selectbox("选择要计算总和的数值列", numeric_columns)

                    summary_data = {
                        "总处理时间": f"{total_time:.2f} 秒",
                        "合并后总行数": f"{len(combined_df)} 行",
                        # f"{selected_numeric_column}总和": f"{combined_df[selected_numeric_column].sum():.2f}"
                    }

                    st.subheader("处理摘要", divider=True)
                    st.table(pd.DataFrame([summary_data]).T.rename_axis(None, axis=1))

                    st.subheader("数据预览", divider=True)
                    st.dataframe(combined_df, use_container_width=True)

                    output = export_file(df=combined_df, selected_columns=selected_columns, mode="combine")
                    st.download_button(
                        label=f"下载文件",
                        data=output.getvalue(),
                        file_name=f"合并后表格.xlsx",
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                else:
                    st.warning("没有成功处理任何文件。请检查上传的文件是否有效。")

    with tab2:
        uploaded_file = st.file_uploader("选择需要拆分的Excel表格", type=["xlsx", "xls"], accept_multiple_files=False)

        st.warning("请确保表格中所有列的列名都不相同", icon="🚨")

        if uploaded_file:
            possible_columns = find_possible_columns(uploaded_file)
            selected_column = st.selectbox("选择目标列", possible_columns)

            if selected_column:
                df = pd.read_excel(uploaded_file, sheet_name=0)

                # Find where the selected_column is in the data
                max_row_index = min(10, len(df))
                data_start_row = None

                for row_index in range(max_row_index):
                    for col_index in range(df.shape[1]):
                        cell_value = df.iloc[row_index, col_index]
                        if cell_value == selected_column:
                            data_start_row = row_index + 1
                            break
                    if data_start_row is not None:
                        break

                if data_start_row is None:
                    st.error(f"未能在文件中找到列 '{selected_column}'")
                else: 
                    # Set the column names
                    column_names = df.iloc[data_start_row - 1]
                    
                   # Ensure unique column names
                    unique_columns = make_unique(column_names)

                    # Assign unique column names to the DataFrame
                    df.columns = unique_columns

                    # Remove the header row (data_start_row - 1) to keep only the data
                    df = df.iloc[data_start_row:].reset_index(drop=True)

                    # Drop rows where all columns are NaN
                    df = df.dropna(how='all')

                    # Get options from the selected_column
                    options = df[selected_column].dropna().unique().tolist()

                    selected_option = st.selectbox("选择要输出的项", options)

                    if selected_option:
                        # Output all rows with that selected option
                        result_df = df[df[selected_column] == selected_option]

                        st.subheader("筛选结果", divider=True)
                        st.dataframe(result_df, use_container_width=True)

                        output = export_file(df=result_df, mode="separate")

                        st.download_button(
                            label=f"下载筛选结果",
                            data=output.getvalue(),
                            file_name=f"筛选结果_{selected_option}.xlsx",
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )

    with tab3:
        st.write("数据稽核工具")
        st.write("coming soon")

    with tab4:
        uploaded_file = st.file_uploader("选择需要可视化的Excel表格", type=["xlsx", "xls"], accept_multiple_files=False)
        
        st.write("请输入表格的标题所在行数(例:标题在第5行, 则输入5)")

        header_row = st.number_input("标题所在行数", min_value=1, max_value=10)

        if uploaded_file and header_row:
            # Load the uploaded Excel file into a pandas DataFrame
            df = pd.read_excel(uploaded_file, engine='openpyxl', header=header_row-1)

            # Display the data summary
            st.header('概况')
            st.write(df.describe())

            # Display the DataFrame
            st.header('预览')
            st.dataframe(df)

            # Display different chart options
            st.header('图表')
            chart_option = st.selectbox('Select chart type', ['Bar Chart', 'Line Chart', 'Area Chart'])

            if chart_option == 'Bar Chart':
                st.bar_chart(df)
            elif chart_option == 'Line Chart':
                st.line_chart(df)
            elif chart_option == 'Area Chart':
                st.area_chart(df)

            # Additional tools for analysis
            st.header('更多工具')
            if st.checkbox('Show correlation matrix'):
                st.write(df.corr())

            if st.checkbox('Show missing values'):
                st.write(df.isnull().sum())

if __name__ == "__main__":
    main()
