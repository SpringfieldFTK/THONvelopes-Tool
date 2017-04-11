import shlex
from collections import deque
import json


class Parameter():
    def __init__(self, name, description, required):
        self.name = name
        self.description = description
        self.required = required


class Command:
    def __init__(self, keywords, description='', apply_to_all=False, commands=None, parameters=None, confirmation=False):
        self.keywords = [x.lower().strip() for x in keywords]
        self.description = description
        self.apply_to_all = apply_to_all
        self.parameters = []
        self.commands = []
        self.entities = []
        self.function = None
        self.confirmation = confirmation

        if commands:
            for command_json in commands:
                command = Command(**command_json)
                self.add_command(command)

        if parameters:
            for parameter_json in parameters:
                parameter = Parameter(**parameter_json)
                self.add_parameter(parameter)

    def matches(self, command):
        return command.lower().strip() in self.keywords

    def add_command(self, command):
        self.commands.append(command)
        self.commands.sort(key=lambda x: x.keywords[0])

    def add_parameter(self, parameter):
        self.parameters.append(parameter)

    def set_entities(self, entities):
        self.entities = entities
        for command in self.commands:
            command.set_entities(entities)

    def print_help(self, full=False, depth=0):
        print(depth*"\t" + "{0}\t-\t{1}".format(self.keywords[0], self.description))
        if full:
            if len(self.commands) == 0:
                print('|'.join(self.keywords), end=' ')
                for parameter in self.parameters:
                    if parameter.required:
                        print("[{0}]".format(parameter.name), end=' ')
                    else:
                        print("({0})".format(parameter.name), end=' ')
                print()
                for parameter in self.parameters:
                    print("\t" + parameter.name, end=' - ')
                    print("required" if parameter.required else "optional", end=': ')
                    print(parameter.description)
            else:
                for command in self.commands:
                    command.print_help(depth=1)

    def set_function(self, function):
        self.function = function

    def call(self, args):
        if len(args) > 0 and args[0] == "help":
            self.print_help(full=True)
            return

        if len(self.commands) > 0:
            if len(args) == 0:
                print("You must enter a command!")
                self.print_help(full=True)
                return
            sub = args.popleft()
            for command in self.commands:
                if command.matches(sub):
                    return command.call(args)
            print("Error command {0} {1} not found".format(self.keywords[0], sub))
            for command in self.commands:
                command.print_help()
        else:
            min_args = sum(1 for p in self.parameters if p.required)
            if len(args) < min_args:
                print("ERROR: Too Few Arguments: {0} arguments expected, {1} received".format(min_args, len(args)))
                self.print_help(full=True)
                return

            if len(self.parameters) < len(args):
                print("ERROR: Too Many Arguments: Max of {0} arguments, {1} received".format(len(self.parameters),
                                                                                             len(args)))
                self.print_help(full=True)
                return

            if self.confirmation:
                response = ""
                while len(response) == 0 or (response[0].lower() != "y" and response[0].lower() != "n"):
                    print("This command has the potential to delete large amounts of data you may not be able to undo this. Continue? [Y/N]")
                    response = input()
                    if len(response) > 0 and response[0].lower() == "n":
                        return
            print("Executing... Please Wait...")
            if self.apply_to_all:
                for entity in self.entities:
                    self.function(entity, args)
            else:
                self.function(args)
            print("Done")

    def get(self, keyword):
        for command in self.commands:
            if command.matches(keyword):
                return command


class Interface:
    def __init__(self):
        self.commands = []
        self.entities = []

    def set_entities(self, list):
        self.entities = list
        for command in self.commands:
            command.set_entities(self.entities)

    def register_command(self, command):
        self.commands.append(command)
        self.commands.sort(key=lambda x: x.keywords[0])
        command.set_entities(self.entities)

    def run(self):
        while True:
            try:
                args = deque(shlex.split(input()))
            except ValueError as error:
                print("Argument Error: {0}".format(error))
                continue
            arg = args.popleft()
            if arg == "help":
                for command in self.commands:
                    command.print_help(full=False)
                continue
            if arg == "exit":
                exit()
            found = False
            for command in self.commands:
                if command.matches(arg):
                    command.call(args)
                    found = True
            if not found:
                print("Error: Command {0} not recognized".format(arg))

    def get(self, keyword):
        for command in self.commands:
            if command.matches(keyword):
                return command

    def parse_json(self, file):
        with open('commands.json') as data_file:
            data = json.load(data_file)

        for json_command in data:
            command = Command(**json_command)
            self.register_command(command)
