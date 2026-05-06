import sys
import os

# Asegurar que el directorio raíz esté en sys.path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from controllers.main_controller import main

if __name__ == '__main__':
    main()
