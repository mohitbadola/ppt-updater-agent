from agno.agent import Agent
from agno.models.mistral import MistralChat
from agno_ppt_excel_agent import (
    ExtractExcelData,
    ExtractPPTText,
    UpdatePPTWithExcel
)
import os
import logging

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def create_sync_agent():
    """Create and configure the synchronization agent"""
    agent = Agent(
        model=MistralChat(id="mistral-small-2506"),
        tools=[
            ExtractExcelData(),
            ExtractPPTText(), 
            UpdatePPTWithExcel()
        ],
        instructions="""
You are an intelligent PowerPoint-Excel synchronization agent. Your job is to:

1. Extract data from Excel files (numbers, text, labels, key-value pairs)
2. Extract text content from PowerPoint presentations
3. Intelligently match and update PowerPoint content with new Excel data
4. Preserve all formatting, layout, and non-data content

WORKFLOW:
1. First, use extract_excel_data() to get all data from the Excel file
2. Then, use extract_ppt_text() to understand the current PowerPoint content
3. Finally, use update_ppt_with_excel() to apply updates intelligently

MATCHING STRATEGIES:
- Direct key-value pairs (label -> value)
- Number similarity matching (handle formatting like $, %, commas)
- Contextual text matching
- Preserve original formatting when updating numbers

Be thorough but conservative - only update content when you're confident it's correct.
Always provide a summary of what was updated.
        """,
        markdown=True,
        show_tool_calls=True
    )
    
    return agent