# 'make' by itself runs the 'all' target
.DEFAULT_GOAL := all

.PHONY: all
all:	heading black yamllint pylint prospector bandit vulture

.PHONY: fmt
fmt:	heading black

.PHONY: lint
lint:	heading yamllint pylint prospector bandit vulture

.PHONY: heading
heading:
	@echo "\033[92m"
	@echo '  ____ _                 __  __       _       _         ____ _               _    '
	@echo ' / ___(_)___  ___ ___   |  \/  | __ _(_)_ __ | |_      / ___| |__   ___  ___| | __'
	@echo '| |   | / __|/ __/ _ \  | |\/| |/ _` | | `_ \| __|____| |   | `_ \ / _ \/ __| |/ /'
	@echo '| |___| \__ \ (_| (_) | | |  | | (_| | | | | | ||_____| |___| | | |  __/ (__|   < '
	@echo ' \____|_|___/\___\___/  |_|  |_|\__,_|_|_| |_|\__|     \____|_| |_|\___|\___|_|\_\'
	@echo ""
	@echo "Static code analyzing with black, yamllint, pylint, prospector, bandit and vulture"
	@echo "\033[0m"

.PHONY: black
black:
	@echo "\033[92m---- Python auto-formating with black (Coding consistancy) --------------------- INFO\033[0m"
	find . -name "*.py" | xargs black --diff --line-length 110
	find . -name "*.py" | xargs black --line-length 110
	@echo ""

.PHONY: yamllint
yamllint:
	@echo "\033[92m---- YAML linting with yamllint (Coding standard) ------------------------------ INFO\033[0m"
	find . -name "*.yaml" | xargs yamllint
	@echo ""

.PHONY: pylint
pylint:
	@echo "\033[92m---- Python linting with pylint (Coding standard and bad code smell) ----------- INFO\033[0m"
	find . -name "*.py" | xargs pylint --rcfile .pylintrc
	@echo ""

.PHONY: prospector
prospector:
	@echo "\033[92m---- Python linting with prospector (Coding standard and bad code smell) ------- INFO\033[0m"
	find . -name "*.py" | xargs prospector --profile .prospector_profile.yaml
	@echo ""

.PHONY: bandit
bandit:
	@echo "\033[92m---- Python linting with bandit (Security) ------------------------------------- INFO\033[0m"
	find . -name "*.py" | xargs bandit
	@echo ""

.PHONY: vulture
vulture:
	@echo "\033[92m---- Python dead code analysis with vulture (Cleaning) ------------------------- INFO\033[0m"
	vulture .
	@echo ""
