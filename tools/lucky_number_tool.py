"""
Lucky Number Tool.

This tool generates a lucky number based on a user's
birth date and the current date.

The implementation is intentionally incomplete and must
be finalized during the laboratory.
"""
from datetime import date
from .tool import Tool


def lucky_number(birth_date):
    """
    Generate a lucky number.

    Parameters:
        birth_date (str): User birth date in format DDMMYYYY.

    Returns:
        int: The generated lucky number.
    """
    today = date.today()
    today_str = today.strftime("%d%m%Y")

    all_digits = today_str + birth_date
    total = sum(int(digit) for digit in all_digits)
    return total


# Create and return the tool definition.
#
# Configure the Tool instance.
#
# Name:
#     A short unique identifier used internally by the agent.
#
# Description:
#     A natural language description explaining
#     when the tool should be used.
#
# Parameters:
#     A dictionary describing all arguments expected
#     by the lucky_number function.
#
# Callback:
#     The Python function that must be executed when
#     the tool is invoked.

lucky_number_tool = Tool(
    name="lucky_number",
    description=(
        "Generates a lucky number based on the user's birth date and today's "
        "date"
    ),
    parameters={
        "type": "object",
        "properties": {
            "birth_date": {
                "type": "string",
                "description": (
                    "The user's birth date in format DDMMYYYY, e.g. 31121993 "
                    "for 31/12/1993"
                )
            }
        },
        "required": ["birth_date"]
    },
    callback=lucky_number
)
