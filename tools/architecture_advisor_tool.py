"""
Architecture Advisor Tool.

This tool recommends a software architecture style
(monolith, modular monolith, or microservices) based on
project scale and budget constraints.
"""

from .tool import Tool


def architecture_advisor(active_users, requests_per_second, budget):
    """
    Recommend an architecture style.

    Parameters:
        active_users (int): Estimated number of active users.
        requests_per_second (int): Estimated peak requests per second.
        budget (str): One of "low", "medium", "high".

    Returns:
        str: Recommendation with justification.
    """
    budget = budget.lower()

    if active_users < 1000 and requests_per_second < 50:
        recommendation = "Monolith"
        reason = (
            "Scara redusă nu justifică overhead-ul operațional al "
            "microserviciilor. Un monolit e mai rapid de dezvoltat și "
            "mentenanță e mult mai simplă."
        )
    elif active_users < 50000 and requests_per_second < 500:
        recommendation = "Modular Monolith"
        reason = (
            "Scara medie beneficiază de o structură modulară bine separată "
            "intern, fără complexitatea rețelei distribuite a "
            "microserviciilor. Permite extragerea ulterioară a modulelor "
            "în servicii separate, dacă e nevoie."
        )
    else:
        recommendation = "Microservices"
        reason = (
            "Scara mare și trafic ridicat justifică separarea în servicii "
            "independente, scalabile individual, cu deployment-uri "
            "izolate."
        )

    if budget == "low" and recommendation == "Microservices":
        reason += (
            " Atenție: bugetul redus poate face dificilă susținerea "
            "costurilor operaționale (orchestrare, monitorizare, echipă "
            "DevOps) specifice microserviciilor."
        )

    return f"Recomandare: {recommendation}. Motiv: {reason}"


architecture_advisor_tool = Tool(
    name="architecture_advisor",
    description=(
        "Recommends a software architecture style (monolith, modular "
        "monolith, or microservices) based on the number of active users, "
        "peak requests per second, and available budget. Use this when "
        "the user asks for architecture recommendations for a project."
    ),
    parameters={
        "type": "object",
        "properties": {
            "active_users": {
                "type": "integer",
                "description": "Estimated number of active users of the system"
            },
            "requests_per_second": {
                "type": "integer",
                "description": "Estimated peak requests per second"
            },
            "budget": {
                "type": "string",
                "description": "Available budget: 'low', 'medium', or 'high'"
            }
        },
        "required": ["active_users", "requests_per_second", "budget"]
    },
    callback=architecture_advisor
)