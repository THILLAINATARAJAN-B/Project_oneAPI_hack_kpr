from collections import OrderedDict
from pc_controller import (
    PCController_files,
    PCController_Application,
    PCController_Basic,
    PCController_MS_App,
    PCController_Win_security,
    PCController_Paint,
    PCController_MS_Store,
    PCController_file_explorer,
    PCController_gmail
)
from pc_controller_online import PCcontoller_online
   # Import the Gmail controller

# Global variable to store recent commands
recent_commands = OrderedDict()

def add_recent_command(command, value, max_size=10):
    """Adds a command to the list of recent commands."""
    if (command, value) in recent_commands:
        recent_commands.move_to_end((command, value))  # Move to end if exists
    else:
        recent_commands[(command, value)] = None
        if len(recent_commands) > max_size:
            recent_commands.popitem(last=False)  # Remove oldest command

def get_recent_commands():
    """Returns a list of recent commands."""
    return list(recent_commands.keys())

def execute_command(command, value, result_label=None, root=None, input_entry=None):
    # Initialize controller instances
    controller_files = PCController_files(root)
    controller_Application = PCController_Application(root)
    controller_Basic = PCController_Basic(root)
    controller_MS_App = PCController_MS_App(root)
    controller_Win_security = PCController_Win_security(root)
    controller_Paint = PCController_Paint(root)
    controller_MS_Store = PCController_MS_Store(root)
    controller_file_explorer = PCController_file_explorer(root)
    controller_gmail = PCController_gmail(root)

    """
    Executes the command based on the parsed result and updates recent commands.
    """
    try:
        # Add the command to the recent commands list
        add_recent_command(command, value)

        # File Explorer-related commands
        if command == 'open_this_pc':
            controller_file_explorer.open_this_pc()
        elif command == 'search_document':
            controller_file_explorer.search_document()
        elif command == 'move_file':
            source, destination = value
            controller_file_explorer.move_file(source, destination)
        elif command == 'cut_file':
            controller_file_explorer.cut_file(value)
        elif command == 'paste_file':
            controller_file_explorer.paste_file(value)
        elif command == 'empty_recycle_bin':
            controller_file_explorer.empty_recycle_bin()
        elif command == 'create_shortcut':
            controller_file_explorer.create_shortcut(value)
        elif command == 'delete_all_files_in_downloads':
            controller_file_explorer.delete_all_files_in_downloads()
        elif command == 'open_documents_folder':
            controller_file_explorer.open_documents_folder()
        elif command == 'open_pictures_folder':
            controller_file_explorer.open_pictures_folder()
        elif command == 'open_music_folder':
            controller_file_explorer.open_music_folder()
        elif command == 'open_downloads_folder':
            controller_file_explorer.open_downloads_folder()
        elif command == 'create_text_file':
            controller_file_explorer.create_text_file()
        elif command == 'check_file_size':
            controller_file_explorer.check_file_size(value)
        elif command == 'restore_file_from_recycle_bin':
            controller_file_explorer.restore_file_from_recycle_bin(value)
        elif command == 'rename_folder_in_documents':
            old_name, new_name = value
            controller_file_explorer.rename_folder_in_documents(old_name, new_name)
        elif command == 'show_hidden_files':
            controller_file_explorer.show_hidden_files()
        elif command == 'hide_file_extensions':
            controller_file_explorer.hide_file_extensions()
        elif command == 'show_file_extensions':
            controller_file_explorer.show_file_extensions()
        elif command == 'compress_folder':
            controller_file_explorer.compress_folder(value)
        elif command == 'extract_zip_file':
            controller_file_explorer.extract_zip_file(value)
        elif command == 'sort_files_by_size':
            controller_file_explorer.sort_files_by_size()
        elif command == 'sort_files_by_date_modified':
            controller_file_explorer.sort_files_by_date_modified()
        elif command == 'select_all_files_in_folder':
            controller_file_explorer.select_all_files_in_folder()
        elif command == 'delete_folder_permanently':
            controller_file_explorer.delete_folder_permanently(value)
        elif command == 'pin_folder_to_quick_access':
            controller_file_explorer.pin_folder_to_quick_access(value)
        elif command == 'unpin_folder_from_quick_access':
            controller_file_explorer.unpin_folder_from_quick_access(value)
        elif command == 'share_file_via_network':
            controller_file_explorer.share_file_via_network(value)
        elif command == 'rename_multiple_files':
            old_name, new_name = value
            controller_file_explorer.rename_multiple_files(old_name, new_name)

        # Paint-related commands
        elif command == 'open_microsoft_paint':
            controller_Paint.open_microsoft_paint()
        elif command == 'create_new_canvas_paint':
            controller_Paint.create_new_canvas_paint()
        elif command == 'open_existing_image_paint':
            controller_Paint.open_existing_image_paint(value)
        elif command == 'save_current_image_paint':
            controller_Paint.save_current_image_paint()
        elif command == 'save_image_as_paint':
            controller_Paint.save_image_as_paint(value)
        elif command == 'close_paint':
            controller_Paint.close_paint()
        elif command == 'resize_canvas_paint':
            width, height = value
            controller_Paint.resize_canvas_paint(width, height)
        elif command == 'undo_last_action_paint':
            controller_Paint.undo_last_action_paint()
        elif command == 'redo_last_action_paint':
            controller_Paint.redo_last_action_paint()
        elif command == 'draw_rectangle_paint':
            start_x, start_y, end_x, end_y = value
            controller_Paint.draw_rectangle_paint(start_x, start_y, end_x, end_y)
        elif command == 'draw_circle_paint':
            center_x, center_y, radius = value
            controller_Paint.draw_circle_paint(center_x, center_y, radius)
        elif command == 'draw_line_paint':
            start_x, start_y, end_x, end_y = value
            controller_Paint.draw_line_paint(start_x, start_y, end_x, end_y)
        elif command == 'draw_freeform_shape_paint':
            controller_Paint.draw_freeform_shape_paint(value)
        elif command == 'change_brush_size_paint':
            controller_Paint.change_brush_size_paint(value)
        elif command == 'change_brush_color_paint':
            controller_Paint.change_brush_color_paint(value)
        elif command == 'change_brush_type_paint':
            controller_Paint.change_brush_type_paint(value)
        elif command == 'fill_shape_with_color_paint':
            x, y, color = value
            controller_Paint.fill_shape_with_color_paint(x, y, color)
        elif command == 'use_eraser_tool_paint':
            start_x, start_y, end_x, end_y = value
            controller_Paint.use_eraser_tool_paint(start_x, start_y, end_x, end_y)
        elif command == 'select_area_paint':
            start_x, start_y, end_x, end_y = value
            controller_Paint.select_area_paint(start_x, start_y, end_x, end_y)
        elif command == 'copy_selected_area_paint':
            controller_Paint.copy_selected_area_paint()
        elif command == 'cut_selected_area_paint':
            controller_Paint.cut_selected_area_paint()
        elif command == 'paste_copied_area_paint':
            x, y = value
            controller_Paint.paste_copied_area_paint(x, y)
        elif command == 'zoom_in_image_paint':
            controller_Paint.zoom_in_image_paint()
        elif command == 'zoom_out_image_paint':
            controller_Paint.zoom_out_image_paint()
        elif command == 'set_image_as_background_paint':
            controller_Paint.set_image_as_background_paint()
        elif command == 'add_text_to_image_paint':
            text, x, y = value
            controller_Paint.add_text_to_image_paint(text, x, y)
        elif command == 'change_font_size_paint':
            controller_Paint.change_font_size_paint(value)
        elif command == 'change_font_color_paint':
            controller_Paint.change_font_color_paint(value)
        elif command == 'draw_polygon_paint':
            controller_Paint.draw_polygon_paint(value)
        elif command == 'change_canvas_background_color_paint':
            controller_Paint.change_canvas_background_color_paint(value)
        elif command == 'delete_selected_area_paint':
            controller_Paint.delete_selected_area_paint()
        elif command == 'move_selected_area_paint':
            dx, dy = value
            controller_Paint.move_selected_area_paint(dx, dy)
        elif command == 'merge_two_images_paint':
            controller_Paint.merge_two_images_paint(value)
        elif command == 'use_color_picker_paint':
            x, y = value
            controller_Paint.use_color_picker_paint(x, y)
        elif command == 'resize_selected_area_paint':
            width, height = value
            controller_Paint.resize_selected_area_paint(width, height)
        elif command == 'export_image_png_paint':
            controller_Paint.export_image_png_paint(value)
        elif command == 'export_image_jpeg_paint':
            controller_Paint.export_image_jpeg_paint(value)
        elif command == 'export_image_bmp_paint':
            controller_Paint.export_image_bmp_paint(value)
        elif command == 'print_image_paint':
            controller_Paint.print_image_paint()
        elif command == 'set_print_options_paint':
            controller_Paint.set_print_options_paint()
        elif command == 'view_image_full_screen_paint':
            controller_Paint.view_image_full_screen_paint()

        # MS Store-related commands
        elif command == 'open_microsoft_store':
            controller_MS_Store.open_microsoft_store()
        elif command == 'search_for_app':
            controller_MS_Store.search_for_app(value)
        elif command == 'install_app':
            controller_MS_Store.install_app(value)
            print("Installing",value)
        elif command == 'update_installed_apps':
            controller_MS_Store.update_installed_apps()
        elif command == 'check_for_app_updates':
            controller_MS_Store.check_for_app_updates()
        elif command == 'enable_auto_update_apps':
            controller_MS_Store.enable_auto_update_apps()
        elif command == 'disable_auto_update_apps':
            controller_MS_Store.disable_auto_update_apps()
        elif command == 'filter_apps_by_category_store':
            controller_MS_Store.filter_apps_by_category_store(value)
        elif command == 'search_for_free_apps':
            controller_MS_Store.search_for_free_apps()
        elif command == 'search_for_paid_apps':
            controller_MS_Store.search_for_paid_apps()
            
        # Gmail-related commands
        elif command == 'open_gmail':
            controller_gmail.open_gmail()
        elif command == 'compose_new_email':
            controller_gmail.compose_new_email()
        elif command == 'send_email':
            controller_gmail.send_email()
        elif command == 'search_email':
            controller_gmail.search_email(value)
        elif command == 'mark_email_read':
            controller_gmail.mark_email_read()
        elif command == 'mark_email_unread':
            controller_gmail.mark_email_unread()
        elif command == 'delete_email':
            controller_gmail.delete_email()
        elif command == 'open_spam_folder':
            controller_gmail.open_spam_folder()
        elif command == 'logout_gmail':
            controller_gmail.logout_gmail()
        # ... (other Gmail commands)

        # Other existing commands
        elif command == 'open_windows_security':
            controller_Win_security.open_windows_security()
        elif command == 'set_brightness':
            controller_Basic.set_brightness(value)
        elif command == 'open_settings':
            controller_Basic.open_settings()
        elif command == 'open':
            if value == 'settings':
                controller_Basic.open('settings')
            elif value == 'my pc':
                controller_files.open('my pc')
            elif value in ['C:', 'D:', 'E:', 'F:']:
                controller_files.open(value)
            else:
                controller_files.open(value)
        elif command == 'open_application':
            controller_Application.open_application(value)
        elif command == 'create_folder':
            controller_files.create_folder(value)
        elif command == 'rename_folder':
            old_name, new_name = value
            controller_files.rename_folder(old_name, new_name)
        elif command == 'delete_file_or_folder':
            controller_files.delete_file_or_folder(value)
        elif command == 'search_file':
            controller_files.search_file(value)
        elif command == 'close_application':
            controller_Application.close_application(value)
        elif command == 'open_task_manager':
            controller_Basic.open_task_manager()
        elif command == 'take_screenshot':
            controller_Basic.take_screenshot(value)
        elif command == 'search':
            controller_Basic.search_file(value)
        elif command == 'refresh':
            controller_Basic.refresh()
        elif command == 'maximize':
            controller_Basic.maximize()
        elif command == 'minimize':
            controller_Basic.minimize()
        elif command == 'close':
            if value == 'settings':
                controller_Basic.close_settings()
            elif value == 'my pc':
                controller_files.close_my_pc()
            elif value in ['C:', 'D:', 'E:', 'F:']:
                controller_files.close(value)
            else:
                controller_files.close_target(value)
        elif command == 'word_content':
            controller_files.create_word_doc_with_content(value, f'{value}.docx')
        elif command == 'create_google_account':
            PCcontoller_online.create_google_account(value)
        else:
            raise ValueError(f"Unknown command: {command}")
    
    except Exception as e:
        if result_label:
            result_label.config(text=f"Error executing '{command}': {e}")
        else:
            print(f"Error executing '{command}': {e}")
    
    finally:
        if root and input_entry:
            root.after(1000, lambda: input_entry.focus_set())