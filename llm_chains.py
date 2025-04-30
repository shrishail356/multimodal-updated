# llm_chains.py

import groq
import json
import time
from langchain.chains import LLMChain
from langchain.chains.retrieval_qa.base import RetrievalQA
from langchain.embeddings import HuggingFaceInstructEmbeddings
from langchain.memory import ConversationBufferWindowMemory
from langchain.prompts import PromptTemplate
from langchain.vectorstores import Chroma
import chromadb
import yaml
from dotenv import load_dotenv
import os
from typing import List, Tuple, Dict, Generator, Optional
import subprocess
import pyautogui
import time
from PIL import ImageGrab
import webbrowser
from langchain_community.tools import DuckDuckGoSearchRun
import re
import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import glob
from pathlib import Path
import platform
from docx import Document
from openpyxl import Workbook
from pptx import Presentation

# Load environment variables and configuration
load_dotenv()
with open("config.yaml", "r") as f:
    config = yaml.safe_load(f)

SYSTEM_PROMPT = """You are an expert AI assistant with advanced reasoning capabilities. 
You must respond in JSON format following this structure:
{
    "title": "Your step title or phase description",
    "content": "Your detailed explanation",
    "next_action": "continue OR final_answer",
    "confidence": "0-100 percentage indicating confidence in this step"
}

Guidelines for your response:
1. Break down complex problems into clear steps
2. Evaluate multiple perspectives and approaches
3. Consider context from conversation history
4. Acknowledge uncertainties and limitations
5. Provide evidence and reasoning for conclusions
6. Quantify confidence levels when possible

If analyzing specific content (images, documents, audio):
- Explicitly reference the content
- Describe relevant features and details
- Connect observations to conclusions

Remember: Always format responses as valid JSON."""

SIMPLE_SYSTEM_PROMPT = """You are an expert AI assistant providing clear and concise responses.
Respond in JSON format:
{
    "title": "Response Title",
    "content": "Your detailed answer",
    "confidence": "0-100 confidence score"
}"""

def web_search(query: str, num_results: int = 3) -> str:
    """
    Perform a web search using DuckDuckGo via LangChain.
    
    Args:
        query: Search query
        num_results: Number of results to return
        
    Returns:
        String containing formatted search results
    """
    try:
        search_tool = DuckDuckGoSearchRun()
        results = search_tool.invoke(query)
        
        if not results:
            return "No web search results found."
            
        return f"Web Search Results:\n{results}"
    
    except Exception as e:
        return f"Web search error: {str(e)}. Proceeding with AI knowledge only."

def minimize_active_window():
    """Minimize the currently active window."""
    try:
        # Use platform-specific hotkey
        if platform.system() == 'Windows':
            pyautogui.hotkey('win', 'down')
        elif platform.system() == 'Darwin':  # macOS
            pyautogui.hotkey('command', 'm')
        else:  # Linux
            pyautogui.hotkey('alt', 'f9')
        return True
    except Exception as e:
        print(f"Error minimizing window: {str(e)}")
        return False

def resize_active_window():
    """Resize the currently active window to a smaller size."""
    try:
        # Get screen dimensions
        screen_width, screen_height = pyautogui.size()
        
        # Calculate new window size (50% of screen)
        new_width = screen_width // 2
        new_height = screen_height // 2
        
        # Calculate position to center the window
        new_x = (screen_width - new_width) // 2
        new_y = (screen_height - new_height) // 2
        
        # Use platform-specific window management
        if platform.system() == 'Windows':
            pyautogui.hotkey('win', 'left')  # Snap to left half
        elif platform.system() == 'Darwin':  # macOS
            pyautogui.hotkey('command', 'option', 'left')  # Move to left half
        else:  # Linux
            pyautogui.hotkey('ctrl', 'alt', 'left')  # Move to left workspace
        
        return True
    except Exception as e:
        print(f"Error resizing window: {str(e)}")
        return False

def resize_browser_window():
    """Resize the active browser window using a more reliable method."""
    try:
        # Get screen dimensions
        screen_width, screen_height = pyautogui.size()
        
        # Calculate new dimensions (65% of screen)
        new_width = int(screen_width * 0.65)
        new_height = int(screen_height * 0.65)
        
        # Calculate position to center the window
        new_x = (screen_width - new_width) // 2
        new_y = (screen_height - new_height) // 2
        
        # Use platform-specific window management
        if platform.system() == 'Windows':
            pyautogui.hotkey('win', 'up')  # Maximize
            time.sleep(0.5)
            pyautogui.hotkey('win', 'down')  # Restore
        elif platform.system() == 'Darwin':  # macOS
            pyautogui.hotkey('command', 'option', 'f')  # Full screen
            time.sleep(0.5)
            pyautogui.hotkey('command', 'option', 'f')  # Exit full screen
        else:  # Linux
            pyautogui.hotkey('alt', 'f10')  # Maximize
            time.sleep(0.5)
            pyautogui.hotkey('alt', 'f5')  # Restore
        
        return True
    except Exception as e:
        print(f"Error resizing window: {str(e)}")
        return False

def create_office_document(app_type: str, content: str = None) -> str:
    """
    Create and write to Microsoft Office documents using cross-platform libraries.
    
    Args:
        app_type: 'word', 'excel', or 'powerpoint'
        content: Content to write in the document
    """
    try:
        # Create a temporary directory for the documents
        temp_dir = os.path.join(os.path.expanduser("~"), "Documents", "temp")
        os.makedirs(temp_dir, exist_ok=True)
        
        if app_type == 'word':
            doc = Document()
            if content:
                doc.add_paragraph(content)
            file_path = os.path.join(temp_dir, "document.docx")
            doc.save(file_path)
            return f"Created Word document at: {file_path}"
            
        elif app_type == 'excel':
            wb = Workbook()
            ws = wb.active
            if content:
                # Split content by lines and write to cells
                for i, line in enumerate(content.split('\n'), 1):
                    ws.cell(row=i, column=1, value=line)
            file_path = os.path.join(temp_dir, "spreadsheet.xlsx")
            wb.save(file_path)
            return f"Created Excel spreadsheet at: {file_path}"
            
        elif app_type == 'powerpoint':
            prs = Presentation()
            slide = prs.slides.add_slide(prs.slide_layouts[1])  # Title and content layout
            if content:
                title = slide.shapes.title
                title.text = "New Slide"
                content_placeholder = slide.placeholders[1]
                content_placeholder.text = content
            file_path = os.path.join(temp_dir, "presentation.pptx")
            prs.save(file_path)
            return f"Created PowerPoint presentation at: {file_path}"
            
    except Exception as e:
        return f"Error creating {app_type} document: {str(e)}"

class GroqLLM:
    """
    Handles interactions with the Groq API, including retry logic and response formatting.
    """
    
    def __init__(self, model_name: str = "llama-3.1-70b-versatile", retry_attempts: int = 3, retry_delay: int = 1):
        """
        Initialize the GroqLLM instance.
        
        Args:
            model_name: Name of the model to use
            retry_attempts: Number of times to retry failed API calls
            retry_delay: Delay in seconds between retries
        """
        self.client = groq.Groq()
        self.model_name = model_name
        self.retry_attempts = retry_attempts
        self.retry_delay = retry_delay

        # Update command patterns to include ChatGPT
        self.command_patterns = [
            # Add ChatGPT specific pattern
            (r'(?i)open\s+chatgpt(?:\s+and\s+search\s+for\s+(.+)|$)', 'chatgpt'),
            
            # YouTube specific pattern
            (r'(?i)(?:play|search|find)\s+(?:video|videos)?\s*["\']?([^"\']+)["\']?\s+(?:on\s+)?youtube', 'youtube'),
            
            # Website specific pattern - captures site and search term
            (r'(?i)open\s+(youtube|google|twitter|facebook|reddit|wikipedia)(?:\s+and\s+search\s+(.+)|$)', 'website'),
            
            # General web search pattern
            (r'(?i)search\s+(?:for\s+)?["\']?([^"\']+)["\']?(?:\s+on\s+(google|bing|duckduckgo))?', 'web_search'),
            
            # Generic URL pattern
            (r'(?i)(?:open|visit|go\sto)\s+(https?://\S+|(?:www\.)?[a-zA-Z0-9-]+\.[a-zA-Z]{2,}(?:/\S*)?)', 'url')
        ]

        # Add ChatGPT to websites config
        self.websites = {
            'chatgpt': {
                'url': 'https://chat.openai.com',
                'search_url': 'https://chat.openai.com',
            },
            'youtube': {
                'url': 'https://www.youtube.com',
                'search_url': 'https://www.youtube.com/results?search_query={}',
            },
            'google': {
                'url': 'https://www.google.com',
                'search_url': 'https://www.google.com/search?q={}',
            },
            'twitter': {
                'url': 'https://twitter.com',
                'search_url': 'https://twitter.com/search?q={}',
            },
            'facebook': {
                'url': 'https://www.facebook.com',
                'search_url': 'https://www.facebook.com/search/top?q={}',
            },
            'reddit': {
                'url': 'https://www.reddit.com',
                'search_url': 'https://www.reddit.com/search/?q={}',
            },
            'wikipedia': {
                'url': 'https://www.wikipedia.org',
                'search_url': 'https://en.wikipedia.org/wiki/Special:Search?search={}',
            }
        }

        # Add delay configuration for automation
        self.automation_delays = {
            'chatgpt_load': 5,  # Seconds to wait for ChatGPT to load
            'input_delay': 1,   # Seconds to wait before typing
            'send_delay': 0.5   # Seconds to wait before pressing enter
        }

        # Add online compiler to websites config
        self.websites['compiler'] = {
            'url': 'https://www.programiz.com/python-programming/online-compiler/',
            'elements': {
                'editor': 'textarea.CodeMirror-line',
                'run_button': 'button.run-button'
            }
        }
        
        # Update command patterns to include complex actions
        self.command_patterns = [
            # Add new pattern for ChatGPT code generation and execution
            (r'(?i)open\s+chatgpt\s+(?:and\s+)?(?:write|generate)\s+(?:some\s+)?code\s+(?:and|then)\s+execute\s+(?:it\s+)?(?:in|using|with)?\s+(?:an\s+)?online\s+compiler', 'chatgpt_code_execution'),
            # ... existing patterns ...
        ]

        # Add new patterns for file/folder operations
        self.command_patterns.extend([
            (r'(?i)find\s+(?:and\s+)?open\s+(?:folder|directory)\s+(?:named\s+)?["\']?([^"\']+)["\']?', 'find_folder'),
            (r'(?i)find\s+(?:and\s+)?open\s+file\s+(?:named\s+)?["\']?([^"\']+)["\']?', 'find_file'),
            (r'(?i)search\s+(?:for\s+)?(?:folder|directory|file)\s+["\']?([^"\']+)["\']?', 'search_item'),
            (r'(?i)open\s+(?:folder|directory|file)\s+["\']?([^"\']+)["\']?', 'open_item')
        ])

        # Add default search paths
        self.search_paths = [
            os.path.normpath(os.path.expanduser("~")),  # User's home directory
            os.path.normpath(os.path.expanduser("~/Documents")),
            os.path.normpath(os.path.expanduser("~/Desktop")),
            os.path.normpath(os.path.expanduser("~/Downloads")),
            os.path.normpath("C:/"),  # System drive
        ]

        # Add Microsoft Office patterns
        self.command_patterns.extend([
            (r'(?i)open\s+(?:microsoft\s+)?word(?:\s+and\s+write\s+(.+)|$)', 'word'),
            (r'(?i)open\s+(?:microsoft\s+)?excel(?:\s+and\s+write\s+(.+)|$)', 'excel'),
            (r'(?i)open\s+(?:microsoft\s+)?powerpoint(?:\s+and\s+write\s+(.+)|$)', 'powerpoint'),
            (r'(?i)create\s+(?:a\s+)?(?:new\s+)?(word|excel|powerpoint)\s+(?:document|file|presentation)(?:\s+with\s+(.+)|$)', 'office_create'),
        ])

    def windows_search(self, query: str, item_type: str = 'both') -> List[str]:
        """
        Search for files or folders using cross-platform methods.
        
        Args:
            query: Search query
            item_type: 'file', 'folder', or 'both'
            
        Returns:
            List of found paths
        """
        try:
            results = []
            for base_path in self.search_paths:
                try:
                    if item_type in ['both', 'file']:
                        # Search for files
                        for file in glob.glob(f"{base_path}/**/*{query}*", recursive=True):
                            if os.path.isfile(file):
                                results.append(file)
                    
                    if item_type in ['both', 'folder']:
                        # Search for folders
                        for folder in glob.glob(f"{base_path}/**/*{query}*", recursive=True):
                            if os.path.isdir(folder):
                                results.append(folder)
                except Exception:
                    continue
            return results
        except Exception as e:
            return []

    def open_item(self, path: str) -> str:
        """
        Open a file or folder using the default application.
        """
        try:
            # Normalize the path before opening
            normalized_path = os.path.normpath(path)
            
            if platform.system() == 'Windows':
                os.startfile(normalized_path)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.run(['open', normalized_path])
            else:  # Linux
                subprocess.run(['xdg-open', normalized_path])
            
            return f"Successfully opened: {normalized_path}"
        except Exception as e:
            return f"Error opening {normalized_path}: {str(e)}"

    def execute_file_operation(self, operation: str, query: str) -> str:
        """
        Execute file/folder operations.
        """
        try:
            if operation == 'find_folder':
                results = self.windows_search(query, 'folder') or self.windows_search(query, 'file')
            elif operation == 'find_file':
                results = self.windows_search(query, 'file')
            else:  # search_item or open_item
                results = self.windows_search(query, 'both')

            if not results:
                return f"No items found matching '{query}'"

            # Format results
            result_text = f"Found {len(results)} items matching '{query}':\n"
            for i, path in enumerate(results, 1):
                result_text += f"{i}. {path}\n"

            # If it's an open operation, open the first result
            if operation in ['open_item', 'find_file', 'find_folder']:
                result_text += "\n" + self.open_item(results[0])

            return result_text

        except Exception as e:
            return f"Error executing file operation: {str(e)}"

    def execute_web_action(self, action_type: str, params: dict) -> str:
        """Execute web-related actions with improved handling."""
        try:
            if action_type == 'chatgpt':
                query = params.get('query', '')
                url = self.websites['chatgpt']['url']
                
                # Open ChatGPT
                webbrowser.open(url)
                
                if query:
                    # Wait for page to load
                    time.sleep(self.automation_delays['chatgpt_load'])
                    
                    # Move to center of screen (where ChatGPT input usually is)
                    screen_width, screen_height = pyautogui.size()
                    pyautogui.moveTo(screen_width // 2, screen_height - 200)
                    pyautogui.click()
                    
                    # Wait before typing
                    time.sleep(self.automation_delays['input_delay'])
                    
                    # Type the query
                    pyautogui.write(query)
                    
                    # Wait before pressing enter
                    time.sleep(self.automation_delays['send_delay'])
                    pyautogui.press('enter')
                    
                    # Wait for response and copy it
                    time.sleep(5)  # Wait for response to appear
                    pyautogui.hotkey('ctrl', 'a')  # Select all
                    time.sleep(0.5)
                    pyautogui.hotkey('ctrl', 'c')  # Copy
                    time.sleep(0.5)
                    
                    # Get the copied text
                    try:
                        copied_text = pyperclip.paste()
                        return f"Opened ChatGPT, sent query: {query}\nResponse copied to clipboard:\n{copied_text}"
                    except:
                        return f"Opened ChatGPT and sent query: {query} (clipboard access failed)"
                    
                return "Opened ChatGPT"

            if action_type == 'youtube':
                query = params['query']
                search_url = self.websites['youtube']['search_url'].format(query.replace(' ', '+'))
                webbrowser.open(search_url)
                return f"Searching YouTube for: {query}"

            elif action_type == 'website':
                site = params['site'].lower()
                query = params.get('query')
                
                if site in self.websites:
                    if query:
                        url = self.websites[site]['search_url'].format(query.replace(' ', '+'))
                    else:
                        url = self.websites[site]['url']
                    webbrowser.open(url)
                    return f"Opening {site.title()}{' and searching for: ' + query if query else ''}"
                return f"Unknown website: {site}"

            elif action_type == 'web_search':
                query = params['query']
                engine = params.get('engine', 'google').lower()
                if engine in self.websites:
                    search_url = self.websites[engine]['search_url'].format(query.replace(' ', '+'))
                else:
                    search_url = self.websites['google']['search_url'].format(query.replace(' ', '+'))
                webbrowser.open(search_url)
                return f"Searching {engine.title()} for: {query}"

            elif action_type == 'url':
                url = params['url']
                if not url.startswith('http'):
                    url = 'https://' + url
                webbrowser.open(url)
                return f"Opened URL: {url}"

            return "Invalid action type"

        except Exception as e:
            return f"Error executing web action: {str(e)}"

    def execute_code_in_compiler(self, code: str) -> str:
        """Execute code in online compiler using Selenium."""
        try:
            # Initialize Chrome in headless mode
            options = webdriver.ChromeOptions()
            options.add_argument('--headless')
            driver = webdriver.Chrome(options=options)
            
            # Open compiler
            driver.get(self.websites['compiler']['url'])
            
            # Wait for editor to load
            editor = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, self.websites['compiler']['elements']['editor']))
            )
            
            # Clear default code and input new code
            driver.execute_script(f"navigator.clipboard.writeText('{code}')")
            editor.send_keys(pyperclip.paste())
            
            # Click run button
            run_button = driver.find_element(By.CSS_SELECTOR, self.websites['compiler']['elements']['run_button'])
            run_button.click()
            
            # Wait for execution and get output
            time.sleep(2)
            output = driver.find_element(By.CSS_SELECTOR, '.output').text
            
            driver.quit()
            return f"Code execution result:\n{output}"
            
        except Exception as e:
            return f"Error executing code: {str(e)}"

    def execute_chatgpt_code_flow(self) -> str:
        """Handle the complete flow of code generation and execution."""
        try:
            # Open ChatGPT
            result = self.execute_web_action('chatgpt', {
                'query': 'Write a simple Python code example that demonstrates list comprehension'
            })
            
            # Extract code from clipboard (assuming ChatGPT response was copied)
            time.sleep(5)  # Wait for ChatGPT to respond
            code = pyperclip.paste()
            
            # Execute the code
            if code:
                return self.execute_code_in_compiler(code)
            return "No code was generated or copied from ChatGPT"
            
        except Exception as e:
            return f"Error in code generation and execution flow: {str(e)}"

    def execute_chatgpt_action(self, query: str = None, resize: bool = True) -> str:
        """Enhanced ChatGPT interaction with improved window management."""
        try:
            # Simulate Ctrl+N to open new tab
            pyautogui.hotkey('ctrl', 'n')
            time.sleep(0.5)  # Wait for new tab
            
            # Open ChatGPT in the new tab
            webbrowser.get().open_new_tab(self.websites['chatgpt']['url'])
            
            # Wait for page load and browser window to be ready
            time.sleep(self.automation_delays['chatgpt_load'])
            
            if query:
                # Move to input area and click
                screen_width, screen_height = pyautogui.size()
                pyautogui.moveTo(screen_width // 2, screen_height - 200)
                pyautogui.click()
                
                # Type query
                time.sleep(self.automation_delays['input_delay'])
                pyautogui.write(query)
                
                # Send query
                time.sleep(self.automation_delays['send_delay'])
                pyautogui.press('enter')
                
                # Wait for response
                time.sleep(3)
                
                # Copy response
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.5)
                pyautogui.hotkey('ctrl', 'c')
                
                # Get copied response
                response = pyperclip.paste()
                
                # Resize browser window if requested
                if resize:
                    time.sleep(0.5)  # Give browser time to fully load
                    resize_browser_window()
                
                return f"ChatGPT Response:\n{response}"
            
            # Resize empty ChatGPT window if requested
            if resize:
                time.sleep(0.5)  # Give browser time to fully load
                resize_browser_window()
            
            return "Opened ChatGPT in new tab"
            
        except Exception as e:
            return f"Error executing ChatGPT action: {str(e)}"

    def interpret_command(self, user_input: str) -> Tuple[bool, str, List[Dict]]:
        """Enhanced command interpretation with improved ChatGPT handling."""
        # Add new ChatGPT patterns
        chatgpt_patterns = [
            (r'(?i)ask\s+chatgpt\s+(?:about\s+)?(.+)', 'chatgpt_query'),
            (r'(?i)(?:get|generate)\s+code\s+from\s+chatgpt\s+for\s+(.+)', 'chatgpt_code'),
            (r'(?i)open\s+chatgpt\s+(?:and\s+)?(?:ask|write)\s+(.+)', 'chatgpt_query')
        ]
        
        # Add ChatGPT patterns to command patterns
        self.command_patterns = chatgpt_patterns + self.command_patterns
        
        # Add new Microsoft Office patterns at the start of the patterns list
        office_patterns = [
            (r'(?i)open\s+(?:microsoft\s+)?word(?:\s+and\s+write\s+(.+)|$)', 'word'),
            (r'(?i)open\s+(?:microsoft\s+)?excel(?:\s+and\s+write\s+(.+)|$)', 'excel'),
            (r'(?i)open\s+(?:microsoft\s+)?powerpoint(?:\s+and\s+write\s+(.+)|$)', 'powerpoint'),
            (r'(?i)create\s+(?:a\s+)?(?:new\s+)?(word|excel|powerpoint)\s+(?:document|file|presentation)(?:\s+with\s+(.+)|$)', 'office_create'),
        ]
        
        # Add office patterns to command patterns
        self.command_patterns = office_patterns + self.command_patterns
        
        for pattern, cmd_type in self.command_patterns:
            match = re.search(pattern, user_input)
            if match:
                if cmd_type in ['chatgpt_query', 'chatgpt_code']:
                    query = match.group(1)
                    if cmd_type == 'chatgpt_code':
                        query = f"Write code for: {query}"
                    return True, cmd_type, [{'type': 'chatgpt', 'query': query}]
                elif cmd_type in ['word', 'excel', 'powerpoint']:
                    content = match.group(1) if match.groups() else None
                    return True, cmd_type, [{'type': 'office', 'app': cmd_type, 'content': content}]
                elif cmd_type == 'office_create':
                    app_type = match.group(1).lower()
                    content = match.group(2) if len(match.groups()) > 1 else None
                    return True, cmd_type, [{'type': 'office', 'app': app_type, 'content': content}]
                    
                # ... rest of the existing command handling ...
                
        return False, '', []

    def execute_command_chain(self, commands: List[Dict]) -> str:
        """Enhanced command execution with Microsoft Office support."""
        try:
            results = []
            for cmd in commands:
                if cmd['type'] == 'chatgpt':
                    result = self.execute_chatgpt_action(cmd.get('query'), resize=True)
                elif cmd['type'] == 'office':
                    result = create_office_document(cmd['app'], cmd.get('content'))
                    results.append(result)
                else:
                    # ... existing command execution code ...
                    result = self.execute_web_action(cmd['type'], cmd)
                results.append(result)
            return '\n'.join(results)
        except Exception as e:
            return f"Error executing command chain: {str(e)}"

    def make_api_call(
        self,
        messages: List[Dict[str, str]],
        max_tokens: int = 5000,  # Increased from 1000
        temperature: float = 0.3,
        stream: bool = False
    ) -> Dict:
        """
        Make an API call to Groq with retry logic.
        
        Args:
            messages: List of message dictionaries
            max_tokens: Maximum tokens in response
            temperature: Temperature for response generation
            stream: Whether to stream the response
            
        Returns:
            Dict containing the response content
        """
        for attempt in range(self.retry_attempts):
            try:
                # Force JSON completion format only for non-streaming responses
                format_config = {"type": "json_object"} if not stream else None
                
                response = self.client.chat.completions.create(
                    model=self.model_name,
                    messages=messages,
                    max_tokens=max_tokens,
                    temperature=temperature,
                    response_format=format_config,
                    stream=stream
                )
                
                if stream:
                    return response
                
                try:
                    # Try to parse JSON response
                    content = response.choices[0].message.content
                    return json.loads(content)
                except json.JSONDecodeError:
                    # If JSON parsing fails, wrap the response in a valid JSON format
                    return {
                        "title": "Response",
                        "content": content,
                        "next_action": "final_answer",
                        "confidence": 50
                    }
                
            except Exception as e:
                if attempt == self.retry_attempts - 1:
                    error_msg = str(e)
                    if "Failed to generate JSON" in error_msg:
                        # Extract the failed generation if available
                        try:
                            error_data = json.loads(error_msg.split("failed_generation': '")[1].rsplit("'}", 1)[0])
                            return error_data
                        except:
                            pass
                    
                    return {
                        "title": "Error",
                        "content": f"API call failed after {self.retry_attempts} attempts: {error_msg}",
                        "next_action": "final_answer",
                        "confidence": 0
                    }
                time.sleep(self.retry_delay)

    def generate_response_with_steps(
        self,
        prompt: str,
        history: Optional[List] = None
    ) -> Generator[Tuple[str, str, float], None, None]:
        """
        Generate a step-by-step response to the prompt with web search integration.
        """
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT}
        ]

        # Check if prompt needs web search
        needs_web_search = bool(re.search(
            r'\b(what|how|who|when|where|why|which|search|find|lookup|latest|current|news|update|today|now)\b',
            prompt.lower()
        ))
        
        if needs_web_search:
            web_results = web_search(prompt)
            formatted_prompt = f"""Web search results for the query:

{web_results}

Based on the above web search results and your knowledge, please analyze this request and respond step-by-step:

{prompt}

Format each step as:
{{
    "title": "Current step description",
    "content": "Detailed explanation",
    "next_action": "continue OR final_answer",
    "confidence": "0-100 confidence score"
}}"""
        else:
            formatted_prompt = f"""Analyze this request and respond step-by-step:

{prompt}

Format each step as:
{{
    "title": "Current step description",
    "content": "Detailed explanation",
    "next_action": "continue OR final_answer",
    "confidence": "0-100 confidence score"
}}"""
        
        if history:
            for msg in history:
                role = "assistant" if msg.type == "ai" else "user"
                messages.append({"role": role, "content": msg.content})
        
        messages.append({"role": "user", "content": formatted_prompt})
        
        step_count = 1
        while True:
            start_time = time.time()
            step_data = self.make_api_call(messages, max_tokens=2000)
            thinking_time = time.time() - start_time
            
            confidence_str = f" (Confidence: {step_data.get('confidence', 'N/A')}%)"
            step_title = step_data['title'] + confidence_str
            
            yield (step_title, step_data['content'], thinking_time)
            messages.append({"role": "assistant", "content": json.dumps(step_data)})
            
            if step_data['next_action'] == 'final_answer':
                break
                
            step_count += 1
        
        final_prompt = """Provide a final, comprehensive answer based on the previous steps.
        Format as JSON:
        {
            "title": "Final Answer",
            "content": "Your complete response",
            "next_action": "final_answer",
            "confidence": "Overall confidence score (0-100)"
        }"""
        
        messages.append({"role": "user", "content": final_prompt})
        final_data = self.make_api_call(messages, max_tokens=5000)
        yield ("Final Answer", final_data['content'], time.time() - start_time)

    def generate_simple_response(
        self,
        prompt: str,
        history: Optional[List] = None,
        max_tokens: int = 5000  # Increased from 500
    ) -> str:
        """
        Generate a simple response without step-by-step reasoning.
        
        Args:
            prompt: User's input prompt
            history: Optional conversation history
            max_tokens: Maximum tokens in response
            
        Returns:
            String containing the response content
        """
        messages = [
            {"role": "system", "content": SIMPLE_SYSTEM_PROMPT}
        ]
        
        if history:
            for msg in history:
                role = "assistant" if msg.type == "ai" else "user"
                messages.append({"role": role, "content": msg.content})
        
        messages.append({"role": "user", "content": prompt})
        
        response = self.make_api_call(messages, max_tokens=max_tokens)
        return response.get('content', 'Error generating response')

def create_embeddings(embeddings_path: str = config["embeddings_path"]) -> HuggingFaceInstructEmbeddings:
    """Create embeddings instance for vector storage."""
    return HuggingFaceInstructEmbeddings(model_name=embeddings_path)

def create_chat_memory(chat_history: List) -> ConversationBufferWindowMemory:
    """Create conversation memory with specified history."""
    return ConversationBufferWindowMemory(memory_key="history", chat_memory=chat_history, k=3)

def load_vectordb(embeddings: HuggingFaceInstructEmbeddings) -> Chroma:
    """Load or create vector database for document storage."""
    persistent_client = chromadb.PersistentClient("chroma_db")
    return Chroma(
        client=persistent_client,
        collection_name="pdfs",
        embedding_function=embeddings
    )

class GroqChatChain:
    """
    Main chat chain implementation using Groq LLM.
    """
    
    def __init__(self, chat_history: List, model_name: str = "llama-3.1-70b-versatile"):
        """
        Initialize chat chain with conversation history.
        
        Args:
            chat_history: List of previous conversation messages
            model_name: Name of the model to use
        """
        self.memory = create_chat_memory(chat_history)
        self.llm = GroqLLM(model_name=model_name)
    
    def run(self, user_input: str) -> str:
        """
        Process user input and return final response.
        
        Args:
            user_input: User's query or prompt
            
        Returns:
            String containing the final response
        """
        final_response = None
        for title, content, _ in self.run_with_steps(user_input):
            if "Final Answer" in title:
                final_response = content
        return final_response
    
    def run_with_steps(
        self,
        user_input: str
    ) -> Generator[Tuple[str, str, float], None, None]:
        """
        Process user input with visible reasoning steps.
        
        Args:
            user_input: User's query or prompt
            
        Yields:
            Tuples of (step_title, step_content, thinking_time)
        """
        yield from self.llm.generate_response_with_steps(
            user_input,
            history=self.memory.chat_memory.messages
        )

class GroqPDFChatChain:
    """
    Specialized chat chain for handling PDF document queries.
    """
    
    def __init__(self, chat_history: List, model_name: str = "llama-3.1-70b-versatile"):
        """
        Initialize PDF chat chain with conversation history.
        
        Args:
            chat_history: List of previous conversation messages
            model_name: Name of the model to use
        """
        self.memory = create_chat_memory(chat_history)
        self.llm = GroqLLM(model_name=model_name)
        self.vector_db = load_vectordb(create_embeddings())
    
    def run(self, user_input: str) -> str:
        """
        Process user input with PDF context and return final response.
        
        Args:
            user_input: User's query about PDF content
            
        Returns:
            String containing the final response
        """
        final_response = None
        for title, content, _ in self.run_with_steps(user_input):
            if "Final Answer" in title:
                final_response = content
        return final_response
    
    def run_with_steps(
        self,
        user_input: str
    ) -> Generator[Tuple[str, str, float], None, None]:
        """
        Process PDF-related query with visible reasoning steps.
        
        Args:
            user_input: User's query about PDF content
            
        Yields:
            Tuples of (step_title, step_content, thinking_time)
        """
        # Retrieve relevant documents
        docs = self.vector_db.similarity_search(user_input, k=3)
        doc_context = "\n\n".join(
            f"Document {i+1}:\n{doc.page_content}" 
            for i, doc in enumerate(docs)
        )
        
        enhanced_prompt = f"""Context from relevant documents:
        {doc_context}
        
        User question: {user_input}
        
        Please analyze the documents and provide a detailed response."""
        
        yield from self.llm.generate_response_with_steps(
            enhanced_prompt,
            history=self.memory.chat_memory.messages
        )

def load_normal_chain(chat_history: List, model_name: str = "llama-3.1-70b-versatile") -> GroqChatChain:
    """Create and return a standard chat chain."""
    return GroqChatChain(chat_history, model_name=model_name)

def load_pdf_chat_chain(chat_history: List, model_name: str = "llama-3.1-70b-versatile") -> GroqPDFChatChain:
    """Create and return a PDF-specialized chat chain."""
    return GroqPDFChatChain(chat_history, model_name=model_name)