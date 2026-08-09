import re
import ast

def add_function_comments(file_path):
    """
    Add descriptive comments to functions that don't have docstrings.
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    lines = content.split('\n')
    new_lines = []
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        # Check if this line is a function definition
        if re.match(r'^def\s+\w+\s*\(', line):
            func_match = re.match(r'^def\s+(\w+)\s*\((.*?)\)', line)
            if func_match:
                func_name = func_match.group(1)
                params = func_match.group(2)
                
                # Add the function definition line
                new_lines.append(line)
                i += 1
                
                # Check if next non-empty line is a docstring
                next_idx = i
                while next_idx < len(lines) and lines[next_idx].strip() == '':
                    next_idx += 1
                
                # Check if there's already a docstring
                has_docstring = False
                if next_idx < len(lines):
                    next_line = lines[next_idx].strip()
                    if next_line.startswith('"""') or next_line.startswith("'''"):
                        has_docstring = True
                
                # If no docstring, add one with basic description
                if not has_docstring and func_name not in ['index', 'admin_login']:
                    indent = len(line) - len(line.lstrip())
                    indent_str = ' ' * (indent + 4)
                    
                    # Generate a simple comment based on function name
                    desc = generate_comment(func_name)
                    new_lines.append(f'{indent_str}"""')
                    new_lines.append(f'{indent_str}{desc}')
                    new_lines.append(f'{indent_str}"""')
                
                continue
        
        new_lines.append(line)
        i += 1
    
    return '\n'.join(new_lines)

def generate_comment(func_name):
    """Generate a simple comment based on function name."""
    # Convert camelCase/snake_case to readable text
    name = re.sub(r'_', ' ', func_name)
    name = re.sub(r'([a-z])([A-Z])', r'\1 \2', name)
    return f"{name.capitalize()}."

if __name__ == '__main__':
    file_path = 'app.py'
    try:
        result = add_function_comments(file_path)
        with open('app_with_comments.py', 'w', encoding='utf-8') as f:
            f.write(result)
        print(f"Successfully created app_with_comments.py")
    except Exception as e:
        print(f"Error: {e}")
