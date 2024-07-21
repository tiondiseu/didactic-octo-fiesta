import datetime
import calendar
import time
import webbrowser
import os
from icalendar import Calendar, Event
from datetime import timedelta

# Fixed variables
company_name = "Company Name"
employee_name = 'John Doe'
hourly_rate = 7.25
current_year = 2024
mileage_rate = 0.25  # This is the compensation rate per mile.
week_starts_on_sunday = True  # Set to False if the week should start on Monday

def create_ics_file(week_number, daily_data):
    cal = Calendar()
    cal.add('prodid', '-//Work Hours Calendar//mxm.dk//')
    cal.add('version', '2.0')

    for day in daily_data:
        event = Event()
        event.add('summary', f'{company_name} | {day["hours"]} hours')
        
        # Create start and end datetime objects
        start_datetime = datetime.datetime.combine(day['date'], day['start_time'])
        end_datetime = datetime.datetime.combine(day['date'], day['end_time'])
        
        # If end_time is earlier than start_time, it means work ended next day
        if end_datetime < start_datetime:
            end_datetime += timedelta(days=1)

        event.add('dtstart', start_datetime)
        event.add('dtend', end_datetime)
        event.add('dtstamp', datetime.datetime.now())

        # Create description, excluding mileage if miles is 0
        description = f"Hours: {day['hours']} at ${hourly_rate} an hour\n"
        if day['miles'] > 0:
            description += f"Miles: {day['miles']} at ${mileage_rate} a mile\n"
        description += f"Total: ${day['total']:.2f}"

        event.add('description', description)

        cal.add_component(event)

    filename = f"W{week_number}_work_hours.ics"
    with open(filename, 'wb') as f:
        f.write(cal.to_ical())
    
    return os.path.abspath(filename)

def get_date_input():
    while True:
        date_str = input("Please insert date of first day of work for the current week (MM/DD/YY): ")
        try:
            return datetime.datetime.strptime(date_str, "%m/%d/%y")
        except ValueError:
            print("Invalid date format. Please use MM/DD/YY.")


def confirm_week(date, current_year, week_starts_on_sunday):
    # Ensure date is a datetime.date object
    if isinstance(date, datetime.datetime):
        date = date.date()

    # Calculate the week number
    if week_starts_on_sunday:
        week_number = (date - datetime.date(current_year, 1, 1)).days // 7 + 1
    else:
        week_number = date.isocalendar()[1]

    # Find the first day of the year
    first_day = datetime.date(current_year, 1, 1)

    # Find the first Sunday of the year if week starts on Sunday
    if week_starts_on_sunday:
        while first_day.weekday() != 6:
            first_day += datetime.timedelta(days=1)

    # Calculate the start date of the given week
    start_date = first_day + datetime.timedelta(weeks=week_number-1)

    # Calculate the end date of the given week
    end_date = start_date + datetime.timedelta(days=6)

    # Format the date range string
    date_range = f"{start_date.strftime('%B %d %A')} {current_year} through {end_date.strftime('%B %d %A')} {current_year}"

    while True:
        confirm = input(f"You are talking about week {week_number}, with the date range of {date_range}. Is that correct? Use Y for yes, and N for no. ").lower()
        if confirm in ['y', 'n']:
            return (week_number, date_range) if confirm == 'y' else (None, None)
        print("Invalid input. Please enter Y or N.")

def get_time_input(prompt):
    while True:
        time_str = input(prompt)
        try:
            return datetime.datetime.strptime(time_str, "%H%M").time()
        except ValueError:
            print("Invalid time format. Please use military time (e.g., 0915).")
    
def convert_hours_to_minutes(number):
    # Check if the number is an integer
    if isinstance(number, int) or number.is_integer():
        return 0

    decimal_part = number % 1 # Get the decimal part
    rounded = round(decimal_part * 4) / 4 # Round to nearest quarter
    minutes = int(rounded * 60) # Convert to minutes
    return minutes

def calculate_hours(start_time, end_time):
    duration = datetime.datetime.combine(datetime.date.min, end_time) - datetime.datetime.combine(datetime.date.min, start_time)
    hours = duration.total_seconds() / 3600
    rounded_hours = round(hours * 4) / 4  # Round to nearest quarter hour
    return rounded_hours

    
def confirming_time(hours, fractional_hours):
    # Separate the integer and decimal parts of hours
    whole_hours = int(hours)
    fractional_hours = round((hours % 1) * 60)  # Convert decimal part to minutes

    while True:
        # Ask for confirmation
        if fractional_hours != 0:
            confirm_message = f"You inputted {whole_hours} hours and {fractional_hours} minutes. Is that correct? (Y/N): "
        else:
            confirm_message = f"You inputted {whole_hours} hours. Is that correct? (Y/N): "

        confirm_hours = input(confirm_message).strip().lower()

        if confirm_hours == "y":
            return True
        elif confirm_hours == "n":
            print("Let's try again.")
            return False
        else:
            print("Invalid input. Please enter 'Y' for yes or 'N' for no.")

def get_mileage():
    while True:
        try:
            miles_input = input("How many miles did you drive in your personal vehicle to the job site? (Enter 'NA' or 'none' if not applicable) ").lower()
            if miles_input in ['na', 'none']:
                return 0  # Return 0 instead of None
            else:
                miles = float(miles_input)
                confirm = input("Did you drive the same amount of miles returning to your point of origin? Use Y for yes, and N for no. ").lower()
                if confirm == 'y':
                    total_miles = miles * 2
                elif confirm == 'n':
                    return_miles_input = input("How many miles did you drive returning to your point of origin? ").lower()
                    if return_miles_input in ['na', 'none']:
                        return_miles = 0
                    else:
                        return_miles = float(return_miles_input)
                    total_miles = miles + return_miles
                else:
                    print("Invalid input. Please enter Y or N.")
                    continue
            
            while True:
                confirm_total = input(f"Okay, you inputted '{total_miles}' total miles. Is that correct? Use Y for yes, and N for no. ").lower()
                if confirm_total == 'y':
                    return total_miles
                elif confirm_total == 'n':
                    print("Let's try again.")
                    break
                else:
                    print("Invalid input. Please enter Y or N.")
        except ValueError:
            print("Invalid input. Please enter a number, NA, or none.")

def get_week_date_range(week_number, current_year):
    if not 1 <= week_number <= 53:
        return "Invalid week number. Please enter a number between 1 and 53."
    
    first_day = datetime.date(current_year, 1, 1) # Get the first day of the year
    
    # Find the first day of the week (Sunday or Monday) of the year
    while first_day.weekday() != (6 if week_starts_on_sunday else 0):
        first_day += datetime.timedelta(days=1)

    start_date = first_day + datetime.timedelta(weeks=week_number-1) # Calculate the start date of the given week
    end_date = start_date + datetime.timedelta(days=6) # Calculate the end date of the given week
    output = f"{start_date.strftime('%B %d %A')} {current_year} through {end_date.strftime('%B %d %A')} {current_year}" # Format the output string
    
    return output

def calculate_daily_total(hours, miles):
    wage = hours * hourly_rate
    mileage_compensation = miles * mileage_rate
    return wage + mileage_compensation

def format_output(week_number, daily_data):
    total_week = sum(day['total'] for day in daily_data)
    return week_number, total_week, daily_data

def create_html(week_number, total_week, daily_data, date_range, employee_name, hourly_rate, mileage_rate):
    html = f"""
    <html>
    <head>
        <style>
            body {{font-family: Arial, sans-serif;}}
            strong {{color: #0066cc;}}
            .week-total {{font-size: large; font-weight: bold; color: red;}}
            .week-dates {{font-size: medium; color: #333;}}
            .daily-total {{font-weight: bold;}}
            .date {{font-size: 1.1em; font-weight: bold;}}
        </style>
    </head>
    <body>
        <h1 class="week-total">SUBJECT: PAYROLL INVOICE Week {week_number} | Hours & Mileage | Employee {employee_name}</h1>
        <h1 class="week-total">Week {week_number} Total = ${total_week:.2f}</h1>
        <p>Date Range: {date_range}</p>
    """
    
    for day in daily_data:
        if day['fractional_hours'] == 0:
            hours_display = f"{int(day['hours'])}"
        else:
            hours_display = f"{int(day['hours'])} and {int(day['fractional_hours'])} minutes"
        
        html += f"""
        <p class="date">{day['date'].strftime('%B %d %A %Y')}</p>
        <p class="daily-total">TOTAL: ${day['total']:.2f}</p>
        <p>{day['start_time'].strftime('%H%M')} to {day['end_time'].strftime('%H%M')} AKA {day['start_time'].strftime('%I:%M %p')} to {day['end_time'].strftime('%I:%M %p')}</p>
        <p>The total amount of hours is {hours_display}.</p>
        <p>{day['hours']}*${hourly_rate}=${day['hours']*hourly_rate:.2f}</p>
        """

        if day['miles'] == 0:
            html += "<p>Mileage calculation is not being considered.</p>"
        else:
            html += f"""
            <p>{day['miles']/2:.0f} miles there and back.</p>
            <p>({day['miles']/2:.0f}*2)/${mileage_rate}=${day['miles']*mileage_rate:.2f}</p>
            <p>${day['hours']*hourly_rate:.2f}+${day['miles']*mileage_rate:.2f} = ${day['total']:.2f}</p>
            """

    html += """
    </body>
    </html>
    """

    with open("invoice.html", "w") as f:
        f.write(html)
    return os.path.abspath("invoice.html")

def main():
    date = get_date_input()
    week_number, date_range = confirm_week(date, current_year, week_starts_on_sunday)



    if not week_number:
        print("Week number not confirmed. Exiting program.")
        return

    daily_data = []
    while True:
        start_time = get_time_input("In military time, what time did you start? ")
        end_time = get_time_input("In military time, what time did you leave? ")
        hours = calculate_hours(start_time, end_time)
        fractional_hours = convert_hours_to_minutes(hours)
        
        confirm_input_hours = confirming_time(hours, fractional_hours)
        print(confirm_input_hours)
        print("Wonderful! You have confirmed the hours. Moving on to the next question. Waiting 1 second...")
        time.sleep(1)

        miles = get_mileage()
        total = calculate_daily_total(hours, miles)

        daily_data.append({
            'date': date,
            'start_time': start_time,
            'end_time': end_time,
            'hours': hours,
            'fractional_hours': fractional_hours,
            'miles': miles,
            'total': total
        })

        next_day = date + datetime.timedelta(days=1)
        work_next_day = input(f"Great, We've finished for the date '{date.strftime('%B %d, %Y')}' did you work the next day '{next_day.strftime('%B %d %A %Y')}'? Use Y for yes, and N for no. ").lower()
        if work_next_day != 'y':
            break
        date = next_day

    week_number, total_week, daily_data = format_output(week_number, daily_data)
    
    html_path = create_html(week_number, total_week, daily_data, date_range,  employee_name, hourly_rate, mileage_rate)
    webbrowser.open('file://' + html_path)

    ics_path = create_ics_file(week_number, daily_data)
    print(f"ICS file created: {ics_path}")

    output = format_output(week_number, daily_data)
    print("Thank you for giving me all of this information. I will compile it and output it.")
    print(output)


if __name__ == "__main__":
    main()