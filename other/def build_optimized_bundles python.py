def build_optimized_bundles():
    import pandas as pd

    # Initialize variables
    max_width = 590
    max_height = 590
    out_row = 2
    dict_items = {}

    # Load input data
    ws_input = pd.read_excel('SO_Input.xlsx')
    ws_output = pd.DataFrame(columns=["Order ID", "Bundle No.", "SKU", "Qty", "SKU Description",
                                       "SKU Width (mm)", "SKU Height (mm)", "SKU Weight (kg)",
                                       "Bundle Total Width", "Bundle Total Height", "Bundle Total Weight", "Note"])

    # Load input to dictionary
    for index, row in ws_input.iterrows():
        order_id = row[0]
        sku = row[1]
        qty = row[3]
        unit_width = row[4]
        unit_height = row[5]
        unit_length = row[6]
        unit_weight = row[7]

        if qty > 0:
            key = f"{order_id}|{sku}|{unit_width}|{unit_height}|{unit_length}|{unit_weight}"
            if key in dict_items:
                dict_items[key] += qty
            else:
                dict_items[key] = qty

    for loop_order_id in get_unique_orders(dict_items):
        bundle_num = 1

        while True:
            current_height = 0
            row_width = 0
            row_height = 0
            max_row_width = 0
            total_weight = 0
            last_sku = ""

            dict_bundle = {}

            # Build bundle
            for key in dict_items.keys():
                if loop_order_id in key and dict_items[key] > 0:
                    parts = key.split("|")
                    if len(parts) < 6:
                        continue

                    sku = parts[1]
                    unit_width = float(parts[2])
                    unit_height = float(parts[3])
                    unit_weight = float(parts[5])
                    qty = dict_items[key]

                    while qty > 0:
                        # Start new row if needed
                        if row_width + unit_width > max_width or (last_sku and sku != last_sku):
                            current_height += row_height
                            max_row_width = max(max_row_width, row_width)
                            if current_height + unit_height > max_height:
                                break
                            row_width = 0
                            row_height = 0

                        # Add SKU
                        dict_bundle[f"{key}|{len(dict_bundle)}"] = [loop_order_id, sku, unit_width, unit_height, unit_weight]
                        row_width += unit_width
                        row_height = max(row_height, unit_height)
                        total_weight += unit_weight
                        qty -= 1
                        dict_items[key] = qty
                        last_sku = sku

                    if qty > 0:
                        break

            current_height += row_height
            max_row_width = max(max_row_width, row_width)

            # Output bundle
            for key in dict_bundle.keys():
                parts = dict_bundle[key]
                if len(parts) >= 5:
                    ws_output.loc[out_row] = [parts[0], bundle_num, parts[1], 1, parts[1], parts[2], parts[3], parts[4], max_row_width, current_height, total_weight, ""]
                    out_row += 1

            bundle_num += 1
            if not any_left(loop_order_id, dict_items):
                break

    print("Bundle Optimization Complete!")

def get_unique_orders(dict_items):
    return set(key.split("|")[0] for key in dict_items.keys())

def any_left(order_id, dict_items):
    return any(order_id in key and qty > 0 for key, qty in dict_items.items())

build_optimized_bundles()