from openpyxl import load_workbook


def get_db(file_DB):
    keys = {"art": str, "name": str, "mult": float, "formuls": str, "coof": float}
    wb = load_workbook(file_DB)
    ws = wb.active
    db_dict = {}
    first_row_flag = True
    for row in ws:
        if first_row_flag:
            first_row_flag = False
            continue
        db_dict[row[0].value] = {k: keys[k](row[n+1].value) for n, k in enumerate(keys)}
    return db_dict


if __name__ == "__main__":
    file_DB = "DB.xlsx"
    db = get_db(file_DB)
    print(db)
