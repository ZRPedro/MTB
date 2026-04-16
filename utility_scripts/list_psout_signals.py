import mhi.psout
import argparse
import sys
import os

def list_signals(psout_file_path, show_multimeters=False):
    """
    Traverses the .psout file and prints a sorted list of signals.
    """
    if not os.path.exists(psout_file_path):
        print(f"Error: File not found at {psout_file_path}")
        return

    all_found_paths = []

    try:
        with mhi.psout.File(psout_file_path) as f:
            run = f.run(0)
            all_trace_strings = [str(t) for t in run.traces()]
            root = f.call('Root/Main')
            
            def get_val(text, key):
                if f"{key}='" in text:
                    return text.split(f"{key}='")[1].split("'")[0]
                return ""

            def traverse(node, current_display_path=''):
                node_str = str(node)
                node_name = get_val(node_str, "Name")

                for t_str in all_trace_strings:
                    t_comp = get_val(t_str, "Component")
                    
                    if t_comp == node_name:
                        sig = get_val(t_str, "Description")
                        if not sig: sig = get_val(t_str, "Name")
                        if not sig: sig = get_val(t_str, "DataName")
                        
                        if not sig or sig.lower() == 'domain':
                            continue
                        
                        sig = sig.split(':')[0]
                        
                        if sig == node_name:
                            full_path = current_display_path
                        else:
                            full_path = f"{current_display_path}\\{sig}" if current_display_path else sig
                        
                        all_found_paths.append(full_path)

                for call in node.calls():
                    child_name = get_val(str(call), "Name")
                    next_path = f"{current_display_path}\\{child_name}" if current_display_path else child_name
                    traverse(call, next_path)

            traverse(root)
    except Exception as e:
        print(f"An error occurred while reading the psout file: {e}")
        return

    # --- Post-Processing ---
    unique_paths = list(set(all_found_paths))
    filtered_paths = []
    
    for p in unique_paths:
        # If show_multimeters is False, hide them unless they are in MTB
        if not show_multimeters:
            if "multimeter" in p.lower() and "MTB" not in p:
                continue
        filtered_paths.append(p)

    main_canvas_signals = sorted([p for p in filtered_paths if "\\" not in p], key=str.lower)
    hierarchical_signals = sorted([p for p in filtered_paths if "\\" in p], key=str.lower)

    # --- Output ---
    print(f"\nSignal Hierarchy for: {psout_file_path}")
    print("=" * 70)
    
    if main_canvas_signals:
        print("\n[ MAIN CANVAS SIGNALS ]")
        for path in main_canvas_signals:
            print(f"  {path}")

    if hierarchical_signals:
        print("\n[ SUB-MODULE SIGNALS ]")
        for path in hierarchical_signals:
            print(f"  {path}")

def main():
    parser = argparse.ArgumentParser(
        description="List and organize all signals from a PSCAD .psout file hierarchically."
    )
    
    # Required positional argument: The file path
    parser.add_argument(
        "path", 
        help="Path to the .psout file"
    )

    # Optional flag: Show multimeters
    parser.add_argument(
        "-m", "--multimeters", 
        action="store_true", 
        help="Include generic multimeter signals in the output (hidden by default)"
    )

    args = parser.parse_args()

    # Run the script with arguments
    list_signals(args.path, show_multimeters=args.multimeters)

if __name__ == "__main__":
    main()