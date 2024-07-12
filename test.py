# import inquirer

# questions = [
#     inquirer.List('option',
#                   message="What do you want to do?",
#                   choices=['Option 1', 'Option 2', 'Option 3', 'Exit'],
#               ),
# ]

# # Function to handle Option 1
# def handle_option_1():
#     print('You chose Option 1')
#     # Add your logic for Option 1 here

# # Function to handle Option 2
# def handle_option_2():
#     print('You chose Option 2')
#     # Add your logic for Option 2 here

# # Function to handle Option 3
# def handle_option_3():
#     print('You chose Option 3')
#     # Add your logic for Option 3 here

# # Main function to run the CLI prompt
# def main():
#     while True:
#         answers = inquirer.prompt(questions)
#         choice = answers['option']
        
#         if choice == 'Option 1':
#             handle_option_1()
#         elif choice == 'Option 2':
#             handle_option_2()
#         elif choice == 'Option 3':
#             handle_option_3()
#         elif choice == 'Exit':
#             print('Exiting the program.')
#             break

# if __name__ == '__main__':
#     main()

# from yaspin import yaspin
# import time

# with yaspin(text="Loading...", color="yellow") as spinner:
#     time.sleep(3)  # Simulate a task by sleeping for 3 seconds
#     spinner.ok("✔")  # Mark the spinner as finished with a success mark


import click

@click.command()
@click.argument('name')
@click.option('--age', default=None, type=int, help='Your age')
@click.option('--color', type=click.Choice(['red', 'green', 'blue']), help='Choose a color')
def greet(name, age, color):
    """A simple CLI app with enhanced appearance using click"""
    click.secho(f'Hello, {name}!', fg='green', bold=True)
    if age:
        click.secho(f'You are {age} years old.', fg='yellow')
    if color:
        click.secho(f'Your favorite color is {color}.', fg=color)

if __name__ == '__main__':
    greet()
