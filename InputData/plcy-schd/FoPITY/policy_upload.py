import ast
import pandas as pd
from pprint import pprint

def convert_csv_to_policy_elements(csv_file):
    """
    CSV 파일을 읽어서 원래의 PolicyElements 튜플 형태로 변환하는 함수.

    csv_file: 변환할 CSV 파일 경로
    반환값: PolicyElements 형태의 튜플 데이터
    """
    df = pd.read_excel(csv_file, sheet_name='csvfile')

    policy_elements = []

    for _, row in df.iterrows():
        policy_name = tuple(row["Policy"].split(" X "))  # "X"로 연결된 정책을 다시 튜플로 변환
        schedules = []

        for col in df.columns[1:]:  # 첫 번째 컬럼(Policy) 제외하고 스케줄 컬럼 처리
            if pd.notna(row[col]):  # NaN 값이 아닌 경우만 처리
                schedule_name = col
                schedule_data = ast.literal_eval(row[col])  # 문자열을 리스트로 변환
                schedules.append((schedule_name, *schedule_data))

        policy_elements.append((policy_name, *schedules))

    return tuple(policy_elements)

# 변환 실행 (예시: CSV 파일 경로를 실제 파일로 변경 필요)
PolicyElements = convert_csv_to_policy_elements("FoPITY-policy-elements-schedule.xlsx")

# 변환된 데이터 출력 (예제 데이터 일부 출력)
for policy in PolicyElements[:15]:
    print(policy)

# 특정 정책요소만 선택하여 확인
def search_policy_elements(keyword, policy_elements):
    matching_policies = [
        policy for policy in policy_elements if any(keyword in item for item in policy[0])
    ]

    if matching_policies:
        pprint(matching_policies)
    else: 
        print(f"{keyword}을 포함하는 정책요소를 찾을수 없습니다.")

search_policy_elements("GRA", PolicyElements)


# Global Constants
# ----------------
FirstYear = 2021 # Update this based on the first simulated year of the model run.
FinalYear = 2100 # This should be the final year supported by EPS.mdl (i.e., 2100), not the final year actually used in a given region.
MaxSchedules = 9
MaxSubscripts = 3
RoundingDigits = 3


# Functions
# ---------

# Write Policy and Subscript Headers
def WritePolicyAndSubscriptHeaders():
  f.write("Policy,")
  for Subscript in range(1,MaxSubscripts+1):
    f.write("Subscript "+str(Subscript)+",")

# Write policy and subscript names, adding commas for unused subscripts
def WritePolicyAndSubscriptNames(PolicyElement):
  for PolicyProperty in range(MaxSubscripts+1):
    # If the policy has no subscripts, just write the policy name plus commas.
    if type(PolicyElement[0]) is str:
      # Only write the name and commas on our first pass through the loop.
      if PolicyProperty == 0:
        f.write(PolicyElement[0]+",,,,")
    elif len(PolicyElement[0])-1 >= PolicyProperty:
      f.write(PolicyElement[0][PolicyProperty]+",")
    else:
      f.write(",")

# Returns the schedule (a list of ordered pairs) to use when writing data for a policy element
def SetActiveSchedule(PolicyElement):
  for ItemIndex in range(len(PolicyElement)):
    # Do not check the policy and subscript names (the first item in PolicyElement)
    if ItemIndex == 0:
      continue
    # Filter non-numerical characters out of the "Schedule X" string in the schedule row.
    ItemScheduleNum = int(list(filter(str.isdigit, PolicyElement[ItemIndex][0]))[0])
    # Check if that number matches the number of the current schedule file we are writing.
    if Schedule == ItemScheduleNum:
      # If we have a match, return all the ordered pairs (all elements after the first) for the matching schedule.
      return PolicyElement[ItemIndex][1:]
  # If we do not have a match, return all the ordered pairs (all elements after the first) for the first listed
  # schedule for that policy element, which is the default schedule.
  return PolicyElement[1][1:]

# Write a .csv file formatted for use by Vensim
def WriteVensimFile():

  # Write header row
  WritePolicyAndSubscriptHeaders()
  for Year in range(FirstYear,FinalYear):
    f.write(str(Year)+",")
  f.write(str(FinalYear)+"\n")
  
  # Write policy element rows
  for PolicyElement in PolicyElements:
  
    WritePolicyAndSubscriptNames(PolicyElement)

    ActiveSchedule = SetActiveSchedule(PolicyElement)
    
    # Write policy implementation percentages for each year
    for Year in range(FirstYear,FinalYear+1):
      # Find the ordered pairs most closely enclosing the active year
      PairBelow = ActiveSchedule[0]
      PairAbove = ActiveSchedule[len(ActiveSchedule)-1]
      for OrderedPair in ActiveSchedule:
        if OrderedPair[0] <= Year:
          PairBelow = OrderedPair
        if OrderedPair[0] >= Year:
          PairAbove = OrderedPair
          break

      # If the enclosing pairs match each other, they also match the active year, so we
      # simply write the implementation percentage from one of the pairs.
      if PairAbove == PairBelow:
        ImplementationPerc = PairBelow[1]
      # Otherwise, we linearly interpolate between the enclosing pairs
      else:
        FractionBetweenYears = (Year-PairBelow[0])/(PairAbove[0]-PairBelow[0])
        ImplementationPerc = PairBelow[1]+FractionBetweenYears*(PairAbove[1]-PairBelow[1])

      # We round the implementation percentage to the correct number of digits and write it
      ImplementationPerc = round(ImplementationPerc, RoundingDigits)

      # Python sometimes writes 1 as "1.0" and 0 as "0.0" when it is a calculated value.
      # It is cleaner to see it as "1" or "0" in the output, so if the value is 1 or 0,
      # we convert the float to an integer.  This doesn't change the value.
      if ImplementationPerc == 1 or ImplementationPerc == 0:
        ImplementationPerc = int(ImplementationPerc)
      
      f.write(str(ImplementationPerc))
      # If this was not the last year, we add a comma
      if Year < FinalYear:
        f.write(",")
        
    # New line for next policy element
    f.write("\n")

# define rounding

def round_to_three_decimal_places(x):
    if isinstance(x, (int, float)):
        return round(x, 3)
    return x

# Write a .csv file formatted for use by the web app
def WriteWebAppFile():

  # Write header row
  WritePolicyAndSubscriptHeaders()
  for Year in range(FirstYear,FinalYear):
    f.write("Year,Imp %,")
  f.write("Year,Imp %\n")

  # Write policy element rows
  for PolicyElement in PolicyElements:
  
    WritePolicyAndSubscriptNames(PolicyElement)
    
    ActiveSchedule = SetActiveSchedule(PolicyElement)

# Write schedule data
    for Year in range(FirstYear, FinalYear + 1):
        if Year - FirstYear < len(ActiveSchedule):
         # Round numerical components to 3 decimal places
            first_element = round_to_three_decimal_places(ActiveSchedule[Year - FirstYear][0])
            second_element = round_to_three_decimal_places(ActiveSchedule[Year - FirstYear][1])

            f.write(str(first_element) + ",")
            f.write(str(second_element) + ",")
        elif Year < FinalYear:
            f.write(",,")
    f.write(",\n")
    

def WritePolicyElementsFile():
  f.write("Policy Element Subscript\n")
  for PolicyElement in PolicyElements:
    # If the policy has no subscripts, just write the policy name followed by " X".
    if type(PolicyElement[0]) is str:
      f.write(PolicyElement[0]+" X\n")
    # Otherwise, write the policy name and subscripts with " X " as delimiter (and no trailing " X")
    else:
      f.write(" X ".join(PolicyElement[0])+"\n")
 
def CheckForScheduleErrors():
  ErrorFound = 0

  f = open("FoPITY-Error-Log.txt", 'w')

  for PolicyElement in PolicyElements:
    for ScheduleNum in range(1,len(PolicyElement)):
      # Extract the ordered pairs (all elements after the first) for the schedule we are checking.
      ScheduleToCheck = PolicyElement[ScheduleNum][1:]

      # Create lists of all years and implementation fractions used in the schedule we are checking.
      Years = []
      ImplementationFractions = []
      for OrderedPair in ScheduleToCheck:
        Years.append(OrderedPair[0])
        ImplementationFractions.append(OrderedPair[1])

      # A set eliminates duplicate values, so we compare the length of the list to its length after
      # converting it to a set.  The lengths are unequal if there are duplicate values in the list.
      if len(set(Years)) != len(Years):
        f.write("Duplicate year(s) found in Schedule "+str(ScheduleNum)+" of Policy Element: "+str(PolicyElement[0])+"\n") 
        ErrorFound = 1
      
      # Ensure all years are between FirstYear and FinalYear inclusive.
      if any(Year < FirstYear or Year > FinalYear for Year in Years):
        f.write("Year(s) prior to FirstYear or after FinalYear found in Schedule "+str(ScheduleNum)+" of Policy Element: "+str(PolicyElement[0])+"\n")
        ErrorFound = 1

      # Ensure all years are in ascending order.
      if (Years != sorted(Years)):
        f.write("Year(s) are not in ascending order in Schedule "+str(ScheduleNum)+" of Policy Element: "+str(PolicyElement[0])+"\n")
        ErrorFound = 1
      
      # Ensure all years are integers.
      if any(not isinstance(Year, int) for Year in Years):
        f.write("Non-integer year(s) found in Schedule "+str(ScheduleNum)+" of Policy Element: "+str(PolicyElement[0])+"\n")
        ErrorFound = 1
      
      # Ensure all implementation fractions are between 0 and 1 inclusive.
      if any(ImplementationFraction < 0 or ImplementationFraction > 1 for ImplementationFraction in ImplementationFractions):
        f.write("Out-of-bounds implementation fraction(s) found in Schedule "+str(ScheduleNum)+" of Policy Element: "+str(PolicyElement[0])+"\n")
        ErrorFound = 1   

  f.close()
  return(ErrorFound)


# Main Program
# ------------

if CheckForScheduleErrors() == 0:

  import os
  if os.path.exists("FoPITY-Error-Log.txt"):
    os.remove("FoPITY-Error-Log.txt")

  for Schedule in range(1,MaxSchedules+1):

    # Begin writing the .csv file for Vensim
    f = open("FoPITY-"+str(Schedule)+".csv", 'w')

    WriteVensimFile()

    # Done writing the .csv file for Vensim
    f.close()

    # Begin writing the .csv file for the web app
    f = open("FoPITY-"+str(Schedule)+"-WebApp.csv", 'w')
    
    WriteWebAppFile()

    # Done writing the .csv file for the web app
    f.close()

  # Write policy elements file
  f = open("FoPITY-policy-elements.csv",'w')

  WritePolicyElementsFile()

  # Done writing the policy elements file
  f.close()
