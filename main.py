import excel_operations as excl_op

def main():
    # Select folder and retun list of excels in folder
    llista_excel = excl_op.listexcel()

    # Unió i format dels excels del banc
    excl_op.combiexcel(llista_excel)


if __name__ == "__main__":
    main()
