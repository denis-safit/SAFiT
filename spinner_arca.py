import sys, time, os

frames = "|/-\\"
i = 0
flag = sys.argv[1] if len(sys.argv) > 1 else ""

while not os.path.exists(flag):
    sys.stdout.write(f"\r  {frames[i % 4]}  Aggiornamento ARCA in corso...  ")
    sys.stdout.flush()
    i += 1
    time.sleep(0.25)

sys.stdout.write("\r  [OK] Aggiornamento ARCA completato!          \n")
sys.stdout.flush()
