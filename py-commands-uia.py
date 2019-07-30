import sys
import datetime
import traceback
from pywinauto.application import Application
from pywinauto import mouse
from pywinauto import win32api


def retrieve_project_parameters():
    
    parameters = sys.argv

    parameters_number = parameters.index("-traces") if "-traces" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        traces = parameters[parameters_number]
    else:
        traces = ""

    parameters_number = parameters.index("-backend") if "-backend" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        backend = parameters[parameters_number]
    else:
        backend = ""

    app_dict = {}

    parameters_number = parameters.index("-app") if "-app" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        app_dict = {"1": parameters[parameters_number]}
    else:
        app_dict = {"1": ""}

    main_window_dict = {}

    parameters_number = parameters.index("-main_window") if "-main_window" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        main_window_dict = {"1": parameters[parameters_number]}
    else:
        main_window_dict = {"1": ""}

    child_window_dict = {}

    parameters_number = parameters.index("-child_window1") if "-child_window1" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        child_window_dict = {"1": parameters[parameters_number]}

    parameters_number = parameters.index("-child_window2") if "-child_window2" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        child_window_dict["2"] = parameters[parameters_number]

    parameters_number = parameters.index("-child_window3") if "-child_window3" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        child_window_dict["3"] = parameters[parameters_number]

    parameters_number = parameters.index("-child_window4") if "-child_window4" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        child_window_dict["4"] = parameters[parameters_number]

    parameters_number = parameters.index("-child_window5") if "-child_window5" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        child_window_dict["5"] = parameters[parameters_number]

    parameters_number = parameters.index("-command") if "-command" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        command = parameters[parameters_number]
    else:
        command = ""

    parameters_number = parameters.index("-value") if "-value" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        value = parameters[parameters_number]
    else:
        value = ""

    parameters_number = parameters.index("-hover") if "-hover" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        hover = parameters[parameters_number]
    else:
        hover = ""

    parameters_number = parameters.index("-wait") if "-wait" in parameters else None
    if parameters_number is not None:
        parameters_number = parameters_number + 1
        wait = parameters[parameters_number]
    else:
        wait = ""
        
    return {
        "traces": traces,
        "backend": backend,
        "app_dict": app_dict,
        "main_window_dict": main_window_dict,
        "child_window_dict": child_window_dict,
        "command": command,
        "value": value,
        "hover": hover,
        "wait": wait,
    }


def validate_project_parameters(parameters):
    
    traces = parameters["traces"]
    backend = parameters["backend"]
    app_dict = parameters["app_dict"]
    main_window_dict = parameters["main_window_dict"]
    child_window_dict = parameters["child_window_dict"]
    command = parameters["command"]
    value = parameters["value"]
    hover = parameters["hover"]
    wait = parameters["wait"]
    
    if traces == "" or traces.upper() == "FALSE":
        traces = False
    elif traces.upper() == "TRUE":
        traces = True
    else:
        return "ERROR: Invalid traces parameter! Parameter = " + str(traces)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Parameters retrieved start * ===")
        
    if backend == "" or backend.upper() == "UIA":
        backend = "uia"
    elif backend.upper() == "WIN32":
        backend = "win32"

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tBackend = " + str(backend))

    if app_dict["1"] == "":
        return "ERROR: Empty app parameter!"

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tApp = " + str(app_dict))

    if main_window_dict["1"] == "":
        return "ERROR: Empty main window parameters!"

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tMain window = " + str(main_window_dict))

    for x in range(0, len(child_window_dict)):
        if child_window_dict[str(x + 1)] == "":
            return "ERROR: Empty child window parameters!"

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tChild window(s) = " + str(child_window_dict))

    if "PRINT" in command.upper():
        command = command.upper()
    elif command.upper() == "CLICK":
        command = command.upper()
    elif command.upper() == "DOUBLECLICK":
        command = command.upper()
    elif command.upper() == "RIGHTCLICK":
        command = command.upper()
    elif command.upper() == "SEND":
        command = command.upper()
    elif command.upper() == "SELECT":
        command = command.upper()
    elif command.upper() == "LOCATION":
        command = command.upper()
    elif command.upper() == "WAIT":
        command = command.upper()
    else:
        return "ERROR: Invalid command parameter! Parameter = " + str(command)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tCommand = " + str(command))

    if "SEND" in command.upper() or "SELECT" in command.upper():
        if value == "":
            return "ERROR: Empty value parameters!"

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tValue = " + str(value))

    if hover == "" or hover.upper() == "FALSE":
        hover = False
    elif hover.upper() == "TRUE":
        hover = True
    else:
        return "ERROR: Invalid hover parameter! Parameter = " + str(hover)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tHover = " + str(hover))

    if command.upper() == "WAIT" and not wait.isdigit():
        return "ERROR: Invalid wait parameter! Parameter = " + str(wait)

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tWait = " + str(wait))

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Parameters retrieved end * ===")
        
    return {
        "traces": traces,
        "backend": backend,
        "app_dict": app_dict,
        "main_window_dict": main_window_dict,
        "child_window_dict": child_window_dict,
        "command": command,
        "value": value,
        "hover": hover,
        "wait": wait,
    }
    

def find_window(backend, obj, level, dict, traces):
    
    success = True

    for loop_x in range(0, len(dict)):
        try:
            temp_list = dict[str(loop_x + 1)].split(",")

            name = None
            process = None
            handle = None
            path = None
            class_name = None
            class_name_re = None
            title = None
            title_re = None
            control_id = None
            control_type = None
            auto_id = None
            framework_id = None

            for loop_y in range(0, len(temp_list)):

                if "name=" in str(temp_list[loop_y]).strip():
                    name = str(temp_list[loop_y]).strip()
                    name = name[5:]

                elif "process=" in str(temp_list[loop_y]).strip():
                    process = str(temp_list[loop_y]).strip()
                    process = process[8:]

                elif "handle=" in str(temp_list[loop_y]).strip():
                    handle = str(temp_list[loop_y]).strip()
                    handle = handle[7:]

                elif "path=" in str(temp_list[loop_y]).strip():
                    path = str(temp_list[loop_y]).strip()
                    path = path[5:]

                elif "class_name=" in str(temp_list[loop_y]).strip():
                    class_name = str(temp_list[loop_y]).strip()
                    class_name = class_name[11:]

                elif "class_name_re=" in str(temp_list[loop_y]).strip():
                    class_name_re = str(temp_list[loop_y]).strip()
                    class_name_re = class_name_re[14:]

                elif "title=" in str(temp_list[loop_y]).strip():
                    title = str(temp_list[loop_y]).strip()
                    title = title[6:]

                elif "title_re=" in str(temp_list[loop_y]).strip():
                    title_re = str(temp_list[loop_y]).strip()
                    title_re = title_re[9:]

                elif "control_id=" in str(temp_list[loop_y]).strip():
                    control_id = str(temp_list[loop_y]).strip()
                    control_id = control_id[11:]

                elif "control_type=" in str(temp_list[loop_y]).strip():
                    control_type = str(temp_list[loop_y]).strip()
                    control_type = control_type[13:]

                elif "auto_id=" in str(temp_list[loop_y]).strip():
                    auto_id = str(temp_list[loop_y]).strip()
                    auto_id = auto_id[8:]

                elif "framework_id=" in str(temp_list[loop_y]).strip():
                    framework_id = str(temp_list[loop_y]).strip()
                    framework_id = framework_id[13:]

            if level == "app":
                if path is not None:
                    obj = Application(backend=backend).connect(path=path)
                elif handle is not None:
                    obj = Application(backend=backend).connect(handle=handle)
                elif process is not None:
                    obj = Application(backend=backend).connect(process=process)
                elif title_re is not None:
                    obj = Application(backend=backend).connect(title_re=title_re)
                else:
                    obj = Application(backend=backend).connect(title=title)

                try:
                    success = obj.is_process_running()
                except:
                    success = False
                    break

            else:
                if level == "main":
                    if name is not None:
                        obj = obj[name]

                    elif backend == "uia":
                        obj = obj.window(class_name=class_name,
                                        class_name_re=class_name_re,
                                        title=title,
                                        title_re=title_re,
                                        control_id=control_id,
                                        control_type=control_type,
                                        auto_id=auto_id,
                                        framework_id=framework_id)
                    else:
                        obj = obj.window(class_name=class_name,
                                        class_name_re=class_name_re,
                                        title=title,
                                        title_re=title_re,
                                        control_id=control_id,
                                        control_type=control_type)

                else:
                    if name is not None:
                        obj = obj[name]

                    elif backend == "uia":
                        obj = obj.child_window(class_name=class_name,
                                            class_name_re=class_name_re,
                                            title=title,
                                            title_re=title_re,
                                            control_id=control_id,
                                            control_type=control_type,
                                            auto_id=auto_id,
                                            framework_id=framework_id)
                    else:
                        obj = obj.child_window(class_name=class_name,
                                            class_name_re=class_name_re,
                                            title=title,
                                            title_re=title_re,
                                            control_id=control_id,
                                            control_type=control_type)

                if traces is True:
                    try:
                        trace_str = level

                        if level is not "main":
                            trace_str = trace_str + " " + str(loop_x + 1)

                        print("")
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\t=== * Locating " + trace_str + " start * ===")
                        print("")
                        obj.print_control_identifiers(depth=1)
                        print("")
                        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\t=== * Locating " + trace_str + " end * ===")
                        print("")

                    except:
                        if level == "main":
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tUNABLE TO LOCATE MAIN WINDOW!")
                        else:
                            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tUNABLE TO LOCATE CHILD WINDOW " + str(loop_x + 1) + "!")
                        success = False
                        break
                    
                if not obj.exists():
                    success = False
                    break

        except:
            print(traceback.format_exc())
            success = False
            break

    if success is True:
        return obj
    else:
        return None


def locate_target(parameters):
    
    traces = parameters["traces"]
    backend = parameters["backend"]
    app_dict = parameters["app_dict"]
    main_window_dict = parameters["main_window_dict"]
    child_window_dict = parameters["child_window_dict"]
    command = parameters["command"]
    value = parameters["value"]
    hover = parameters["hover"]
    wait = parameters["wait"]
    
    app_obj = None
    window_obj = None
    
    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Locate the target start * ===")
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tLocating the main application...")

    app_obj = find_window(backend, "", "app", app_dict, traces)

    if app_obj is None:
        return app_obj, window_obj, "ERROR: Main application not located!"
    else:
        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tMain application located!")

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tLocating the main window...")

    window_obj = find_window(backend, app_obj, "main", main_window_dict, traces)

    if window_obj is None:
        return app_obj, window_obj, "ERROR: Main window not located!"
    else:
        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tMain window located!")

    if not len(child_window_dict) == 0:
        if traces is True:
            print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tLocating the child window(s)...")

        window_obj = find_window(backend, window_obj, "child", child_window_dict, traces)

        if window_obj is None:
            return app_obj, window_obj, "ERROR: Child window(s) not located!"
        else:
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tChild window(s) located!")

    if traces is True:
        print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Locate the target end * ===")
        
    return app_obj, window_obj, "SUCCESS"


def execute_command(parameters, app_obj, window_obj):
    
    traces = parameters["traces"]
    backend = parameters["backend"]
    app_dict = parameters["app_dict"]
    main_window_dict = parameters["main_window_dict"]
    child_window_dict = parameters["child_window_dict"]
    command = parameters["command"]
    value = parameters["value"]
    hover = parameters["hover"]
    wait = parameters["wait"]
    
    try:
        if "PRINT" in command.upper():
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Print control identifiers start * ===")

            if "FILE" in command.upper():
                if value is not "":
                    value = r"C:\print_control_identifiers.txt"

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tPrinting control identifiers to file: " + str(value))

                window_obj.print_control_identifiers(filename=str(value))

            elif value is not "" and value.isdigit():
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tPrinting control identifiers to output variable with depth: " + str(value))

                window_obj.print_control_identifiers(depth=int(value))

            else:
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tPrinting control identifiers to output variable...")

                window_obj.print_control_identifiers()

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tPrinting control identifiers complete!")
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Print control identifiers end * ===")

        elif command.upper() == "WAIT":
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Waiting for control start * ===")

            success = True

            try:
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tWaiting for the control to exist for " + str(wait) + " seconds...")

                window_obj.wait("exists", timeout=int(wait), retry_interval=0.25)

            except:
                success = False

            if success is False:
                return "Error: Element did not appear!"

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tControl exists!")
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Waiting for control end * ===")

        elif command.upper() == "LOCATION":
            command = "LOCATION"

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Print location start * ===")

            window_obj.print_control_identifiers(depth=1)

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Print location end * ===")

        elif "CLICK" in command.upper():
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform click command start * ===")

            mouse_location_x, mouse_location_y = win32api.GetCursorPos()

            if command.upper() == "CLICK":
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to click...")

                window_obj.click_input()

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tClick complete!")

            elif command.upper() == "DOUBLECLICK":
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to double-click...")

                window_obj.click_input(double=True)

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tDouble-click complete!")

            elif command.upper() == "RIGHTCLICK":
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to right-click...")

                window_obj.click_input(button="right")

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tRight-click complete!")

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform click command end * ===")
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Move mouse back to original position start * ===")

            if hover is False:
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to move mouse back to position...")

                mouse.move((mouse_location_x, mouse_location_y))

                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tMoving mouse back to position complete!")

            else:
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tHovering mouse activated, therefore, do NOT move mouse back in position")

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Move mouse back to original position end * ===")

        elif "SEND" in command.upper():
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform send command start * ===")
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to write values: " + str(value))

            window_obj.type_keys(value, with_spaces=True, with_tabs=True, with_newlines=True)

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tWriting values complete!")
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform send command end * ===")

        elif "SELECT" in command.upper():
            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform select command start * ===")

            if value.startswith("item=") and value[5:].isdigit():
                if traces is True:
                    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tValue is digit, therefore, convert to integer")

                value = int(int(value[5:]) - 1)

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tAttempting to select item: " + str(value))

            window_obj.select(value)

            if traces is True:
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "\tSelecting value complete!")
                print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ": " + "=== * Perform select command end * ===")
    
        return True
                
    except:
        print(traceback.format_exc())
        return "ERROR: Unexpected issue!"


def main():
    
    parameters = retrieve_project_parameters()
    
    parameters = validate_project_parameters(parameters)
    if not isinstance(parameters, dict):
        print(str(parameters))
        return
    
    app_obj, window_obj, valid = locate_target(parameters)
    if "ERROR" in valid:
        print(str(valid))
        return
    
    valid = execute_command(parameters, app_obj, window_obj)
    if not valid is True:
        print(str(valid))
        return
    
    print("SUCCESS")
    
    
if __name__ == "__main__":
    main()
