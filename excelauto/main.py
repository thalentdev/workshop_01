from helpers import add_tabular_data, add_chart


def main():
    add_tabular_data()
    add_chart()

    print('DONE: Excel was created in reports folder.')


# if directly on app.py not as module, will run below script
if __name__ == '__main__':
    main()
