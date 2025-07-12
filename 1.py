import pandas as pd
import warnings
warnings.simplefilter(action="ignore",category=FutureWarning)
df = pd.read_excel("Salary_Data.xlsx")
print(df)

new_person = pd.Series([32, "Female", "Bachelor's Software", "Engineer",5,90000],
    index=["Age", "Gender", "Education Level", "Job Title", "Years of Experience", "Salary"])
add_df = df._append(new_person, ignore_index = True)
print(add_df)
