#!/bin/bash
# Git Assistant Script
# Version: 1.2.0
# Author: Florent ALBANY (FAL) - f.albany@serma.com
# License: Proprietary. All rights reserved.
# Description: Interactive tool for common Git operations
#
# Usage:
#   Interactive mode: ./git_assistant.sh
#   Direct command:   ./git_assistant.sh <command-number>
#
# Command Reference:
#   1 - Pull Changes       : Download updates from remote repository
#   2 - Push Changes       : Upload committed changes to remote repository
#   3 - Initialize Repo    : Create new Git repository in current directory
#   4 - Show Status        : Display branch info, changes and recent commits
#   5 - Commit Changes     : Stage and commit changes with message
#   6 - Initialize Submodules: Setup and update all nested repositories
#   0 - Exit               : Quit the script
#
# Copyright (c) 2023 SERMA Group. Unauthorized use prohibited.

# --- Color Definitions ---
C_RESET='\033[0m'
C_RED='\033[0;31m'
C_GREEN='\033[0;32m'
C_YELLOW='\033[0;33m'
C_BLUE='\033[0;34m'
C_CYAN='\033[0;36m'
C_MAGENTA='\033[0;35m'

# --- Helper Functions ---
is_git_repo() {
    git rev-parse --is-inside-work-tree &>/dev/null
    return $?
}

print_header() {
    echo -e "${C_BLUE}----------------------------------------${C_RESET}"
    echo -e "${C_CYAN}$1${C_RESET}"
    echo -e "${C_BLUE}----------------------------------------${C_RESET}"
}

print_info() {
    echo -e "${C_YELLOW}INFO: $1${C_RESET}"
}

print_success() {
    echo -e "${C_GREEN}SUCCESS: $1${C_RESET}"
}

print_error() {
    echo -e "${C_RED}ERROR: $1${C_RESET}"
}

ask_confirm() {
    while true; do
        read -p "Do you want to continue? (y/n): " confirm
        case "$confirm" in
            [Yy]*) return 0 ;;
            [Nn]*) return 1 ;;
            *) echo -e "${C_YELLOW}Please enter Y/y or N/n${C_RESET}" ;;
        esac
    done
}

# --- Git Operations ---
git_pull_changes() {
    print_header "Pull Changes from Remote"
    if ! is_git_repo; then
        print_error "Not a Git repository"
        return 1
    fi
    
    print_info "Downloads latest changes from the remote server and merges them into your current branch"
    echo -e "Operation: ${C_CYAN}git pull${C_RESET}\n"
    
    if ! ask_confirm; then
        echo -e "Pull cancelled.\n"
        return
    fi

    print_info "Updating repository..."
    if git pull; then
        print_success "Repository successfully updated with remote changes"
    else
        print_error "Failed to pull changes. Check your network connection and repository permissions"
    fi
}

git_push_changes() {
    print_header "Push Changes to Remote"
    if ! is_git_repo; then
        print_error "Not a Git repository"
        return 1
    fi
    
    print_info "Uploads your committed changes to the shared remote repository"
    
    echo -e "\n${C_MAGENTA}--- Repository Status ---${C_RESET}"
    git status -sb
    echo -e "${C_MAGENTA}-------------------------${C_RESET}\n"
    
    echo -e "Operation: ${C_CYAN}git push${C_RESET}"
    if ! ask_confirm; then
        echo "Push cancelled."
        return
    fi

    print_info "Uploading changes..."
    if git push; then
        print_success "Changes successfully shared with remote repository"
    else
        print_error "Push failed. You may need to pull recent changes first or check access rights"
    fi
}

git_init_repo() {
    print_header "Initialize New Repository"
    
    if is_git_repo; then
        print_error "A Git repository already exists in this directory"
        return 1
    fi

    print_info "Creates a new Git version control repository in the current folder"
    echo -e "Operation: ${C_CYAN}git init${C_RESET}\n"
    
    if ! ask_confirm; then
        echo "Initialization cancelled."
        return
    fi

    if git init; then
        print_success "New repository successfully created"
    else
        print_error "Failed to initialize repository. Check directory permissions"
    fi
}

git_show_status() {
    print_header "Repository Status Overview"
    
    if ! is_git_repo; then
        print_error "Not a Git repository"
        return 1
    fi
    
    print_info "Shows current branch information, local changes, and recent commit history"
    
    echo -e "\n${C_MAGENTA}--- Branch Information ---${C_RESET}"
    git branch -vv
    
    echo -e "\n${C_MAGENTA}--- Local Changes ---${C_RESET}"
    git status -sb
    
    echo -e "\n${C_MAGENTA}--- Recent Commits (last 5) ---${C_RESET}"
    git log --oneline -n 5
    
    echo -e "${C_MAGENTA}-----------------------${C_RESET}"
}

git_commit_changes() {
    print_header "Commit Changes"
    
    if ! is_git_repo; then
        print_error "Not a Git repository"
        return 1
    fi
    
    print_info "Records changes to the repository with a descriptive message"
    
    echo -e "\n${C_MAGENTA}--- Changes to be committed ---${C_RESET}"
    git diff --cached --stat
    echo -e "${C_MAGENTA}-------------------------------${C_RESET}\n"
    
    # Check if there are staged changes
    if [[ -z $(git diff --cached --name-only) ]]; then
        print_info "No changes staged for commit. Stage all changes now?"
        if ask_confirm; then
            git add .
            print_success "All changes staged for commit"
        else
            echo "Commit cancelled."
            return
        fi
    fi
    
    # Get commit message
    echo
    read -p "Enter commit message (Ctrl+C to cancel): " commit_msg
    if [[ -z "$commit_msg" ]]; then
        print_error "Commit message cannot be empty"
        return 1
    fi
    
    if git commit -m "$commit_msg"; then
        print_success "Changes successfully committed"
        echo -e "\n${C_MAGENTA}--- Commit Details ---${C_RESET}"
        git show --stat
        echo -e "${C_MAGENTA}----------------------${C_RESET}"
    else
        print_error "Commit failed. Check for merge conflicts or empty commit"
    fi
}

git_init_submodules() {
    print_header "Initialize Submodules"
    
    if ! is_git_repo; then
        print_error "Not a Git repository"
        return 1
    fi
    
    print_info "Initializes and clones all nested submodule repositories"
    echo -e "Operation: ${C_CYAN}git submodule update --init --recursive${C_RESET}\n"
    
    if ! ask_confirm; then
        echo "Submodule initialization cancelled."
        return
    fi

    print_info "Initializing submodules..."
    if git submodule update --init --recursive; then
        print_success "All submodules initialized successfully"
        
        # Show submodule status
        echo -e "\n${C_MAGENTA}--- Submodule Status ---${C_RESET}"
        git submodule status
        echo -e "${C_MAGENTA}------------------------${C_RESET}"
    else
        print_error "Submodule initialization failed. Check network access and submodule paths"
    fi
}

# --- Main Program ---
main() {
    # Parse command-line argument if exists
    local choice
    if [[ $# -gt 0 ]]; then
        choice="$1"
    else
        print_header "Git Assistant v1.2.0"
        echo -e "Available commands:"
        echo "1. Pull Changes       : Download updates from remote"
        echo "2. Push Changes       : Upload changes to remote"
        echo "3. Initialize Repo    : Create new repository"
        echo "4. Show Status        : Display repository state"
        echo "5. Commit Changes     : Record changes with message"
        echo "6. Initialize Submodules: Setup nested repositories"
        echo "0. Exit"
        echo
        read -p "Enter command number (0-6): " choice
    fi

    # Perform selected operation
    case "$choice" in
        1) git_pull_changes ;;
        2) git_push_changes ;;
        3) git_init_repo ;;
        4) git_show_status ;;
        5) git_commit_changes ;;
        6) git_init_submodules ;;
        0) print_info "Exiting Git Assistant"; exit 0 ;;
        *) 
            print_error "Invalid option: $choice"
            echo -e "Valid options:"
            echo "1 - Pull Changes       4 - Show Status"
            echo "2 - Push Changes       5 - Commit Changes"
            echo "3 - Initialize Repo    6 - Initialize Submodules"
            echo "0 - Exit"
            exit 1
            ;;
    esac
	    
	# Pause before exit in interactive mode
    if [[ $# -eq 0 ]]; then
        echo
        read -p "Operation complete. Press Enter to exit..."
    fi
}

main "$@"