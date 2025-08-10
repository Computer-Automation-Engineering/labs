#!/usr/bin/env python3

def load_git_instructions():
    """Load git instructions from the text file"""
    try:
        with open('git_instructions.txt', 'r') as file:
            return file.read()
    except FileNotFoundError:
        print("Error: git_instructions.txt not found!")
        return None

def show_categories():
    """Display available categories"""
    categories = [
        "BASIC SETUP",
        "REPOSITORY CREATION", 
        "BASIC WORKFLOW",
        "VIEWING CHANGES",
        "BRANCHING",
        "REMOTE REPOSITORIES",
        "UNDOING CHANGES",
        "USEFUL SHORTCUTS"
    ]
    
    print("\n=== AVAILABLE CATEGORIES ===")
    for i, category in enumerate(categories, 1):
        print(f"{i}. {category}")
    print("9. View All Instructions")
    print("0. Exit")
    return categories

def show_category(content, category_name):
    """Show specific category content"""
    start_marker = f"=== {category_name} ==="
    lines = content.split('\n')
    
    in_category = False
    category_content = []
    
    for line in lines:
        if line.strip() == start_marker:
            in_category = True
            category_content.append(line)
            continue
        elif line.startswith("=== ") and in_category:
            break
        elif in_category:
            category_content.append(line)
    
    if category_content:
        print(f"\n{start_marker}")
        for line in category_content:
            print(line)
    else:
        print(f"Category '{category_name}' not found!")

def main():
    """Main program loop"""
    print("=== GIT INSTRUCTIONS BROWSER ===")
    print("Simple tool to browse Git commands by category")
    
    content = load_git_instructions()
    if content is None:
        return
    
    while True:
        categories = show_categories()
        
        try:
            choice = input("\nEnter your choice (0-9): ").strip()
            
            if choice == '0':
                print("Goodbye!")
                break
            elif choice == '9':
                print(content)
            elif choice.isdigit() and 1 <= int(choice) <= len(categories):
                category_name = categories[int(choice) - 1]
                show_category(content, category_name)
            else:
                print("Invalid choice! Please enter a number between 0-9.")
                
        except KeyboardInterrupt:
            print("\nGoodbye!")
            break
        except Exception as e:
            print(f"Error: {e}")

if __name__ == "__main__":
    main()