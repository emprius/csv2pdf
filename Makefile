.PHONY: install run clean check-tkinter install-tkinter check-pip install-pip create-venv activate-venv

VENV_DIR := venv

install: check-pip check-tkinter create-venv
	$(VENV_DIR)/bin/pip install --upgrade pip
	$(VENV_DIR)/bin/pip install -r requirements.txt

run: install
	$(VENV_DIR)/bin/python main.py

clean:
	rm -rf __pycache__ $(VENV_DIR)

check-tkinter:
	@python3 -c "import tkinter" 2>/dev/null || $(MAKE) install-tkinter

install-tkinter:
	@if [ -f "/etc/arch-release" ]; then \
		echo "Detected Arch-based system. Installing tkinter with pacman..."; \
		sudo pacman -S --noconfirm tk; \
	elif [ -f "/etc/debian_version" ]; then \
		echo "Detected Debian-based system. Installing tkinter with apt..."; \
		sudo apt update && sudo apt install -y python3-tk; \
	else \
		echo "Unsupported operating system. Please install tkinter manually."; \
		exit 1; \
	fi

check-pip:
	@which pip >/dev/null 2>&1 || $(MAKE) install-pip

install-pip:
	@if [ -f "/etc/arch-release" ]; then \
		echo "Detected Arch-based system. Installing pip with pacman..."; \
		sudo pacman -S --noconfirm python-pip; \
	elif [ -f "/etc/debian_version" ]; then \
		echo "Detected Debian-based system. Installing pip with apt..."; \
		sudo apt update && sudo apt install -y python3-pip; \
	else \
		echo "Unsupported operating system. Please install pip manually."; \
		exit 1; \
	fi

create-venv:
	@if [ ! -d "$(VENV_DIR)" ]; then \
		echo "Creating virtual environment..."; \
		python3 -m venv $(VENV_DIR); \
	fi
