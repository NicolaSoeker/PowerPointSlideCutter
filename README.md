# PowerPointSlideCutter
Cuts Power Point many Slides into Seperate Peresntations. Given there is a power point with multiple slides in it. We want to have a powerpoint for each slide containing only that slide. 

We need a place to store test presentations. I suggest a folder called /tests

We need a place to store the code. I suggest a filder called /src

How to set up the dev env:
  initally: 
    - After installing all required packages in your virtual environment, run: 
    
      pip freeze > requirements.txt
  
  second setup: 
      
      python -m venv myenv  # Create a virtual environment
    
      source myenv/bin/activate  # Activate it (Linux/macOS)
      
      myenv\Scripts\activate  # Activate it (Windows)
  
      pip install -r requirements.txt  # Install dependencies





Things we can think about: 
- Code Quality & Formatting: Black (formatter), Flake8 (linter), Pre-commit hooks

