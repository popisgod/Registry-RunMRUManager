import winreg
import click

PATH = winreg.HKEY_CURRENT_USER
RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
RESERVED = 0  

def find_missing_char(char_list):
    '''
    Function to find missing char in RunMRU list.
    '''
    # Convert the list of characters to a set for faster lookup
    char_set = set(char_list)
    
    # Iterate through the alphabet to find the first missing character
    for char in 'abcdefghijklmnopqrstuvwxyz':
        if char not in char_set:
            return char
    
    return None  # Return None if all characters are used

def get_run_key():
    """
    Function to get the registry key.
    """
    try:
        return winreg.OpenKey(PATH, RUN_KEY_PATH, RESERVED, winreg.KEY_ALL_ACCESS)
    except FileNotFoundError:
        click.echo("Error: Registry key not found.")
        return None
    except PermissionError:
        click.echo("Error: Permission denied. Try running the program as an administrator.")
        return None
    except Exception as e:
        click.echo(f"Error: {e}")
        return None

def get_run_history():
    """
    Function to retrieve the list of values from the registry key.
    """
    run_key = get_run_key()
    if run_key is not None:
        try:
            run_pairs = []
            number_of_values = winreg.QueryInfoKey(run_key)[1]
            for i in range(number_of_values):
                run_pairs.append(winreg.EnumValue(run_key, i))
            return run_pairs
        except Exception as e:
            click.echo(f"Error: {e}")
            return []
    else:
        return []

@click.group()
def program():
    """
    Click group for the command-line interface.
    """
    pass

@click.command()
def history():
    """
    Click command to display the run history values.
    """
    run_history = get_run_history()
    if run_history:
        click.echo("Run History:")
        for value_name, value_data, value_type in run_history:
            value_data = value_data.split('\\')[0]
            if value_name != 'MRUList' and value_name != '':
                click.echo(value_data)

@click.command()
@click.argument('value_name', type=click.STRING)
def add(value_name : str):
    """
    Click command to add a new value to the registry key 'RunMRU'.
    """
    run_key = get_run_key()
    if run_key is not None:
        try:
            MRUList = ([pair[1] for pair in get_run_history() if pair[0] == 'MRUList'][0])
            value_key = find_missing_char(list(MRUList))
            
            MRUList = value_key + MRUList 
            winreg.SetValueEx(run_key, 'MRUList' , RESERVED,  winreg.REG_SZ,  MRUList)
            
            winreg.SetValueEx(run_key, value_key , RESERVED,  winreg.REG_SZ,  value_name + r'\1')
            click.echo(f"Added '{value_name}' to the registry.")
        except Exception as e:
            click.echo(f"Error: {e}")

@click.command()
@click.argument('value_name', type=click.STRING)
def delete(value_name : str):
    """
    Click command to delete a specific value from the registry key 'RunMRU'.
    """
    run_key = get_run_key()
    if run_key is not None:
        try:
            run_pairs = get_run_history()
            for value_key, value, _ in run_pairs:
                if value.split('\\')[0] == value_name:
                    winreg.DeleteValue(run_key, value_key)
                    click.echo(f"Deleted value '{value_name}' from the registry.")
                    
                    MRUList = list([pair[1] for pair in run_pairs if pair[0] == 'MRUList'][0])
                    MRUList.remove(value_key)
                    MRUList = ''.join(MRUList)
                    winreg.SetValueEx(run_key, 'MRUList' , RESERVED,  winreg.REG_SZ,  MRUList)
                    break
            else:
                click.echo(f"Value '{value_name}' does not exist in the registry.")
        except Exception as e:
            click.echo(f"Error: {e}")

if __name__ == '__main__':
    program.add_command(history)
    program.add_command(add)
    program.add_command(delete)
    
    program()
