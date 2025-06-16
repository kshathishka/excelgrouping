import streamlit as st
import pandas as pd
import io

def read_excel_file(uploaded_file, sheet_name_input, is_people_file):
    """
    Reads an Excel or CSV file and returns its data as a list of dictionaries,
    the primary column name, and the auto-detected college column name.
    """
    try:
        file_name = uploaded_file.name
        is_csv = file_name.lower().endswith('.csv')

        if is_csv:
            df = pd.read_csv(uploaded_file)
        else:
            try:
                df = pd.read_excel(uploaded_file, sheet_name=sheet_name_input)
            except ValueError:
                st.error(f"Sheet '{sheet_name_input}' not found in {file_name}. Please check the sheet name.")
                return [], '', ''

        if df.empty:
            st.error(f"No data found in {file_name}.")
            return [], '', ''

        # Convert DataFrame to list of dictionaries (records)
        # Ensure all values are converted to string to avoid issues with mixed types, especially for identifiers
        json_data = df.astype(str).to_dict(orient='records')

        # Auto-detect primary column name: Use the first column header found
        primary_column_name = df.columns[0]
        if not primary_column_name:
            st.error(f"Could not detect primary column name in {file_name}. Make sure the sheet has headers.")
            return [], '', ''

        # Auto-detect college column
        auto_detected_college_col = ''
        college_keywords = ['college', 'university', 'institution', 'school', 'uni']

        for header in df.columns:
            lower_header = str(header).lower()
            if any(keyword in lower_header for keyword in college_keywords):
                # Check if this column actually contains non-empty values for at least some rows
                if df[header].dropna().empty:
                    continue # Skip if column is entirely empty after dropping NA
                
                # Check if the column is not just numeric (e.g., student IDs)
                if pd.api.types.is_numeric_dtype(df[header]) and not df[header].astype(str).str.contains('[a-zA-Z]').any():
                    continue # Skip if it looks like only numeric IDs

                auto_detected_college_col = header
                break  # Found a suitable college column, take the first one

        message_type = "People" if is_people_file else "Team Heads"
        st.info(f"Loaded {len(json_data)} {message_type} from '{file_name}' using primary column **'{primary_column_name}'**.")
        if auto_detected_college_col:
            st.info(f"Auto-detected college column: **'{auto_detected_college_col}'**.")
        else:
            st.warning(f"No college column auto-detected for {message_type} file. Grouping will be general if no college columns are found in both files.")

        return json_data, primary_column_name, auto_detected_college_col

    except Exception as e:
        st.error(f"Error processing {uploaded_file.name}: {e}")
        return [], '', ''

st.set_page_config(layout="centered", page_title="Excel/CSV Grouping Tool")

st.title("Excel/CSV Grouping Tool")

st.markdown("""
Upload two Excel or CSV files: one with a list of people and another with team heads.
The tool will group people under team heads. It will attempt a college-based grouping if
it can auto-detect a 'college' column in both files. Otherwise, it will perform a general,
even distribution.
""")

st.write("---")

## Upload People File
st.subheader("1. Upload People Excel/CSV Sheet")
people_file = st.file_uploader("Upload People File (e.g., 1200 people)", type=["xlsx", "xls", "csv"], key="people_file")
people_sheet_name = st.text_input("People Sheet Name (e.g., Sheet1) - Ignored for CSV", value="Sheet1", key="people_sheet")
st.markdown("<p style='font-size: small; color: gray;'>For Excel, specify sheet name. For CSV, the entire file is treated as one sheet. The first column will be used for people's names. College column will be auto-detected.</p>", unsafe_allow_html=True)

people_data = []
people_primary_column_used = ''
people_auto_detected_college_column = ''

if people_file:
    people_data, people_primary_column_used, people_auto_detected_college_column = read_excel_file(people_file, people_sheet_name, True)

st.write("---")

## Upload Team Heads File
st.subheader("2. Upload Team Heads Excel/CSV Sheet")
heads_file = st.file_uploader("Upload Team Heads File (e.g., 80 heads)", type=["xlsx", "xls", "csv"], key="heads_file")
heads_sheet_name = st.text_input("Heads Sheet Name (e.g., Sheet1) - Ignored for CSV", value="Sheet1", key="heads_sheet")
st.markdown("<p style='font-size: small; color: gray;'>For Excel, specify sheet name. For CSV, the entire file is treated as one sheet. The first column will be used for team head names. College column will be auto-detected.</p>", unsafe_allow_html=True)

heads_data = []
heads_primary_column_used = ''
heads_auto_detected_college_column = ''

if heads_file:
    heads_data, heads_primary_column_used, heads_auto_detected_college_column = read_excel_file(heads_file, heads_sheet_name, False)

st.write("---")

## Process Button
if st.button("Process and Group", type="primary"):
    if not people_data:
        st.error('Please upload the People Excel/CSV sheet and ensure data is loaded.')
    elif not heads_data:
        st.error('Please upload the Team Heads Excel/CSV sheet and ensure data is loaded.')
    else:
        grouped_results = []
        assigned_people_identifiers = set()

        def get_identifier(item, primary_column):
            return item.get(primary_column) # Use .get to handle cases where column might be missing

        st.subheader("Grouping Results")

        # College-based grouping logic
        if people_auto_detected_college_column and heads_auto_detected_college_column:
            st.info("Attempting college-based grouping...")

            people_by_college = {}
            for person in people_data:
                college = person.get(people_auto_detected_college_column)
                if college and str(college).strip() != '':
                    normalized_college = str(college).strip().lower()
                    if normalized_college not in people_by_college:
                        people_by_college[normalized_college] = []
                    people_by_college[normalized_college].append(person)

            heads_by_college = {}
            for head in heads_data:
                college = head.get(heads_auto_detected_college_column)
                if college and str(college).strip() != '':
                    normalized_college = str(college).strip().lower()
                    if normalized_college not in heads_by_college:
                        heads_by_college[normalized_college] = []
                    heads_by_college[normalized_college].append(head)

            processed_heads_identifiers = set() # To track heads that have been processed

            for college_name, heads_in_college in heads_by_college.items():
                people_in_college = people_by_college.get(college_name, [])

                if heads_in_college and people_in_college:
                    num_people_in_college = len(people_in_college)
                    num_heads_in_college = len(heads_in_college)
                    base_per_head = num_people_in_college // num_heads_in_college
                    remainder = num_people_in_college % num_heads_in_college

                    current_people_index_in_college = 0
                    for head in heads_in_college:
                        head_id = get_identifier(head, heads_primary_column_used)
                        if head_id is not None:
                            processed_heads_identifiers.add(head_id)

                        count_for_this_head = base_per_head
                        if remainder > 0:
                            count_for_this_head += 1
                            remainder -= 1

                        if count_for_this_head == 0:
                            row = {}
                            for key, value in head.items():
                                row[f"Team Head - {key}"] = value
                            row['Group Member - Status'] = f"No members assigned from {college_name}"
                            grouped_results.append(row)
                        else:
                            for _ in range(count_for_this_head):
                                if current_people_index_in_college < num_people_in_college:
                                    member = people_in_college[current_people_index_in_college]
                                    member_id = get_identifier(member, people_primary_column_used)

                                    if member_id is not None and member_id not in assigned_people_identifiers:
                                        row = {}
                                        for key, value in head.items():
                                            row[f"Team Head - {key}"] = value
                                        for key, value in member.items():
                                            row[f"Group Member - {key}"] = value
                                        grouped_results.append(row)
                                        assigned_people_identifiers.add(member_id)
                                    current_people_index_in_college += 1
                                else:
                                    break
                elif heads_in_college: # Heads in this college but no people from this college
                    for head in heads_in_college:
                        head_id = get_identifier(head, heads_primary_column_used)
                        if head_id is not None:
                            processed_heads_identifiers.add(head_id)
                        row = {}
                        for key, value in head.items():
                            row[f"Team Head - {key}"] = value
                        row['Group Member - Status'] = f"No members from {college_name} assigned to this head"
                        grouped_results.append(row)
            
            # Handle heads who were not processed because their college had no people or no college info
            for head in heads_data:
                head_id = get_identifier(head, heads_primary_column_used)
                if head_id is not None and head_id not in processed_heads_identifiers:
                    row = {}
                    for key, value in head.items():
                        row[f"Team Head - {key}"] = value
                    row['Group Member - Status'] = "No members assigned (No college match or college info)"
                    grouped_results.append(row)

            # Handle people who were not assigned (either no college info, or no head in their college, or couldn't be evenly distributed within college)
            unassigned_people = [person for person in people_data if get_identifier(person, people_primary_column_used) not in assigned_people_identifiers]

            if unassigned_people:
                st.warning(f"{len(unassigned_people)} people could not be assigned based on college grouping. Attempting to assign them generally.")
                
                # Create a list of heads who still have capacity (or all heads if some are empty)
                available_heads = [head for head in heads_data if get_identifier(head, heads_primary_column_used) not in processed_heads_identifiers]
                if not available_heads: # If all heads processed, cycle through all heads again for remaining people
                    available_heads = heads_data

                if available_heads:
                    head_index = 0
                    for person in unassigned_people:
                        if head_index >= len(available_heads):
                            head_index = 0 # Cycle back to the beginning of available heads
                        
                        head_to_assign = available_heads[head_index]
                        row = {}
                        for key, value in head_to_assign.items():
                            row[f"Team Head - {key}"] = value
                        for key, value in person.items():
                            row[f"Group Member - {key}"] = value
                        grouped_results.append(row)
                        head_index += 1
                else:
                    st.warning("No heads available for general assignment of remaining people.")
                    for person in unassigned_people:
                        row = {'Team Head - Status': 'UNASSIGNED (No matching college head or no available head)'}
                        for key, value in person.items():
                            row[f"Group Member - {key}"] = value
                        grouped_results.append(row)

        else:
            st.info('No matching college columns detected or missing in one of the files, performing general grouping.')
            num_people = len(people_data)
            num_heads = len(heads_data)

            if num_heads == 0:
                st.error('Error: No team heads found. Cannot create groups.')
                st.stop()

            base_per_head = num_people // num_heads
            remainder = num_people % num_heads

            people_index = 0

            for i in range(num_heads):
                head = heads_data[i]
                count_for_this_head = base_per_head
                if remainder > 0:
                    count_for_this_head += 1
                    remainder -= 1

                if count_for_this_head == 0:
                    row = {}
                    for key, value in head.items():
                        row[f"Team Head - {key}"] = value
                    row['Group Member - Status'] = 'No members assigned'
                    grouped_results.append(row)
                else:
                    for j in range(count_for_this_head):
                        if people_index < num_people:
                            member = people_data[people_index]
                            row = {}

                            for key, value in head.items():
                                row[f"Team Head - {key}"] = value

                            for key, value in member.items():
                                row[f"Group Member - {key}"] = value
                            grouped_results.append(row)
                            people_index += 1
                        else:
                            break
            
            # Add any remaining unassigned people
            while people_index < num_people:
                person = people_data[people_index]
                row = {'Team Head - Status': 'UNASSIGNED (No more heads available)'}
                for key, value in person.items():
                    row[f"Group Member - {key}"] = value
                grouped_results.append(row)
                people_index += 1


        if not grouped_results:
            st.warning('No groups were formed. Check your input data.')
        else:
            # Create a DataFrame from the grouped results
            output_df = pd.DataFrame(grouped_results)

            # Generate Excel file in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                output_df.to_excel(writer, sheet_name="Grouped Teams", index=False)
                # The 'with' statement handles saving and closing automatically.
                # No need for writer.save()
            output.seek(0) # Rewind to the beginning of the stream

            st.success('Grouping complete! Download your Excel file below.')
            st.download_button(
                label="Download Grouped Teams Excel",
                data=output,
                file_name="grouped_teams.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_button"
            )