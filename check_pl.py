import sys
import os
sys.path.append(os.getcwd())

from src.models.pl_cubo_db import PlCuboDB

print("PL Periods:", PlCuboDB.get_available_periods("Pacifico"))
