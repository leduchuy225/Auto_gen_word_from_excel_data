import re
from datetime import datetime
from dateutil.relativedelta import relativedelta


def replace_placeholder_preserve_style(paragraph, placeholder, replacement):
    buffer = ""
    runs_to_replace = []
    found = False

    # Step 1: Combine runs to find where the placeholder appears
    for run in paragraph.runs:
        buffer += run.text
        runs_to_replace.append(run)

        if placeholder in buffer:
            found = True
            break

    if not found:
        return  # Placeholder not found, do nothing

    # Step 2: Compute pre + replacement + post
    combined_text = "".join(run.text for run in runs_to_replace)
    new_text = combined_text.replace(placeholder, replacement)

    # Step 3: Clear old runs
    for run in runs_to_replace:
        run.text = ""

    # Step 4: Create new runs for text (optional: break into multiple styled runs)
    # We'll just reuse the style from the first matching run
    if runs_to_replace:
        styled_run = runs_to_replace[0]
        styled_run.text = new_text


def evaluate_placeholder(expr: str):
    match = re.match(r"{{\s*(\w+)\((\w+)\)\s*}}", expr)
    if not match:
        return None

    func_name, arg_name = match.groups()

    return func_name, arg_name


def generate_ky_han_tra_no_goc(start_date, total_duration_months, amount, cycle_months):
    start_date = datetime.strptime(start_date, "%d/%m/%Y")

    print(start_date)
    print(total_duration_months)
    print(amount)
    print(cycle_months)

    # Date list
    date_list = []
    elapsed_months = 0
    while elapsed_months <= total_duration_months - cycle_months:
        new_date = start_date + relativedelta(months=elapsed_months)
        date_list.append(new_date)
        elapsed_months += cycle_months
        if elapsed_months == total_duration_months:
            break
        if elapsed_months > total_duration_months - cycle_months:
            new_date = start_date + relativedelta(
                months=total_duration_months - cycle_months
            )
            date_list.append(new_date)

    # Money list
    cycle_count = total_duration_months // cycle_months
    money_each_cycle = amount * cycle_months / total_duration_months
    money_list = [money_each_cycle] * cycle_count
    if (spare_money := amount - (money_each_cycle * cycle_count)) > 0:
        money_list.append(spare_money)

    print(date_list)
    print(money_list)

    # Combine two lists to get result
    result_lines = []

    for index, date in enumerate(date_list):
        amount_str = f"{(money_list[index]):,.0f}".replace(",", ".")
        result_lines.append(
            f"\t- Ngày {date.strftime('%d/%m/%Y')}, số tiền {amount_str} đồng."
        )

    # Combine all into one string
    final_output = "\n".join(result_lines)

    return final_output
