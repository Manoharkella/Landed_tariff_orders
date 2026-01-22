import sqlite3
import os

def view_terminal_format():
    db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tariff_orders.db")
    if not os.path.exists(db_path):
        print(f"Database not found at {db_path}")
        return

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Get all columns
    cursor.execute("PRAGMA table_info(tariff_data)")
    col_info = cursor.fetchall()
    columns = [row[1] for row in col_info if row[1] not in ('id', 'updated_at')]
    
    cursor.execute(f"SELECT {', '.join(columns)} FROM tariff_data")
    rows = cursor.fetchall()

    # User's requested display order and labels
    display_map = [
        ('financial_year', 'financial_year'),
        ('state', 'States'),
        ('discom', 'DISCOM'),
        ('ists_loss', 'ISTS Loss'),
        ('insts_loss', 'InSTS Loss'),
        ('wheeling_loss_11kv', 'Wheeling Loss – 11 kV'),
        ('wheeling_loss_33kv', 'Wheeling Loss – 33 kV'),
        ('wheeling_loss_66kv', 'Wheeling Loss – 66 kV'),
        ('wheeling_loss_132kv', 'Wheeling Loss – 132 kV'),
        ('ists_charges', 'ISTS Charges'),
        ('insts_charges', 'InSTS Charges'),
        ('wheeling_charges_11kv', 'Wheeling Charges – 11 kV'),
        ('wheeling_charges_33kv', 'Wheeling Charges – 33 kV'),
        ('wheeling_charges_66kv', 'Wheeling Charges – 66 kV'),
        ('wheeling_charges_132kv', 'Wheeling Charges – 132 kV'),
        ('css_charges_11kv', 'Cross Subsidy Surcharge – 11 kV'),
        ('css_charges_33kv', 'Cross Subsidy Surcharge – 33 kV'),
        ('css_charges_66kv', 'Cross Subsidy Surcharge – 66 kV'),
        ('css_charges_132kv', 'Cross Subsidy Surcharge – 132 kV'),
        ('css_charges_220kv', 'Cross Subsidy Surcharge – 220 kV'),
        ('additional_surcharge', 'Additional Surcharge'),
        ('electricity_duty', 'Electricity Duty'),
        ('tax_on_sale', 'Tax on Sale'),
        ('fixed_charge_11kv', 'Fixed Charge – 11 kV'),
        ('fixed_charge_33kv', 'Fixed Charge – 33 kV'),
        ('fixed_charge_66kv', 'Fixed Charge – 66 kV'),
        ('fixed_charge_132kv', 'Fixed Charge – 132 kV'),
        ('fixed_charge_220kv', 'Fixed Charge – 220 kV'),
        ('energy_charge_11kv', 'Energy Charge – 11 kV'),
        ('energy_charge_33kv', 'Energy Charge – 33 kV'),
        ('energy_charge_66kv', 'Energy Charge – 66 kV'),
        ('energy_charge_132kv', 'Energy Charge – 132 kV'),
        ('energy_charge_220kv', 'Energy Charge – 220 kV'),
        ('fuel_surcharge', 'Fuel Surcharge'),
        ('tod_charges', 'TOD Charges'),
        ('pf_rebate', 'Power Factor Adjustment Rebate'),
        ('lf_incentive', 'Load Factor Incentive'),
        ('grid_support_parallel_op_charges', 'Grid Support / Parallel Operation Charges'),
        ('ht_ehv_rebate_33_66kv', 'HT / EHV Rebate at 33/66 kV'),
        ('ht_ehv_rebate_132_above', 'HT / EHV Rebate at 132 kV and above'),
        ('bulk_rebate', 'Bulk Consumption Rebate')
    ]

    for row in rows:
        data = dict(zip(columns, row))
        print("="*80)
        print(f"DATABASE RECORD: {data.get('state')} - {data.get('discom')}")
        print("="*80)
        
        for db_key, label in display_map:
            val = data.get(db_key, "NA")
            print(f"{label:<45} : {val}")
        
        print("\n")

    conn.close()

if __name__ == "__main__":
    view_terminal_format()
