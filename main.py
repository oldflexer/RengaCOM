import sys
import win32com.client


def save_in_file(file_name, content):
    with open(f"C:\\Users\\Prioritet\\PyProjects\\RengaCOM\\{file_name}", 'w') as file:
        file.write(content)


def main():
    app = win32com.client.Dispatch("Renga.Application.1")
    app.Visible = True

    if app.OpenProject("D:\\Programs\\Renga Professional\\Projects\\MainProject.rnp") != 0:
        print("Error opening project")
        sys.exit(1)

    project = app.Project
    model = project.Model

    opening_type = "{fc443d5a-b76c-45e5-b91c-520ef0896109}".upper()
    window_type = "{2b02b353-2ca5-4566-88bb-917ea8460174}".upper()
    door_type = "{1cfba99c-01e7-4078-ae1a-3e2ff0673599}".upper()
    wall_type = "{4329112a-6b65-48d9-9da8-abf1f8f36327}".upper()
    room_type = "{f1a805ff-573d-f46b-ffba-57f4bccaa6ed}".upper()

    opening_count = 0
    window_count = 0
    door_count = 0
    wall_count = 0
    room_count = 0

    # Creating and starting an operation before editing the project
    operation = project.CreateOperation()
    operation.Start()

    object_collection = model.GetObjects()
    # print(object_collection.Count)
    for index in range(object_collection.Count):
        obj = object_collection.GetByIndex(index)
        # print(obj.Name, obj.ObjectTypeS)
        # NOTE: using the S-property
        if obj.ObjectTypeS == opening_type:
            opening_count += 1
        if obj.ObjectTypeS == window_type:
            window_count += 1
        if obj.ObjectTypeS == door_type:
            door_count += 1
        if obj.ObjectTypeS == wall_type:
            wall_count += 1
        if obj.ObjectTypeS == room_type:
            room_count += 1

    print(room_count)
    # print((opening_count, window_count, door_count))
    # save_in_file('log.txt', f"Openings: {opening_count} Windows: {window_count} Doors: {door_count}")

    # Applying the changes
    operation.Apply()

    # Closing the project without saving, dumping all the hard work
    app.CloseProject(True)
    app.Quit()


if __name__ == '__main__':
    main()
