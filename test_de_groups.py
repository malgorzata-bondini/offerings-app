from generator_core import get_de_company_and_ldap
import pandas as pd

# Create a mock original row with "Subscribed by Company" value
original_row = pd.Series({
    "Subscribed by Company": "Original Company Value",
    "Support group": "Some Other Group",
    "Name (Child Service Offering lvl 1)": "Test Offering"
})

# Test the function with different support groups
test_groups = [
    ("HS DE IT Service Desk HC", "HS DE"),
    ("HS DE IT Service Desk - MCC", "HS DE"), 
    ("DS DE IT Service Desk -Labs", "DS DE"),
    ("DS DE IT Service Desk - Labs", "DS DE"),  # With space before dash
    ("HS DE IT Service Desk", "HS DE"),
    ("DS DE IT Service Desk", "DS DE"),
    ("Other Support Group", "HS DE"),
    ("Another Group", "DS DE")
]

print("Testing get_de_company_and_ldap function:")
print("=" * 60)

for support_group, receiver in test_groups:
    company, ldap = get_de_company_and_ldap(support_group, receiver, original_row)
    print(f"Support Group: '{support_group}'")
    print(f"Receiver: '{receiver}'")
    print(f"Company: '{company}'")
    print(f"LDAP: '{ldap}'")
    print("-" * 40)
