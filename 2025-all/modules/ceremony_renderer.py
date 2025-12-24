"""Jinja template rendering for ceremony scripts."""

import logging
import re
from typing import Dict, Set
from jinja2 import Environment, FileSystemLoader, TemplateNotFound

logger = logging.getLogger("ceremony_generator")


class CeremonyRenderer:
    """Handles Jinja template rendering for ceremony scripts."""
    
    def __init__(self, template_dir: str):
        self.template_dir = template_dir
        self.env = Environment(loader=FileSystemLoader(template_dir))
    
    def load_template(self, template_filename: str) -> 'jinja2.Template':
        """Load a Jinja template file.
        
        Args:
            template_filename: Name of template file
            
        Returns:
            Jinja2 Template object
            
        Raises:
            TemplateNotFound: If template file doesn't exist
        """
        try:
            template = self.env.get_template(template_filename)
            logger.info(f"Loaded template: {template_filename}")
            return template
        except TemplateNotFound:
            logger.error(f"Template not found: {template_filename}")
            raise
    
    def extract_template_variables(self, template_filename: str) -> Set[str]:
        """Extract all {{ variable }} tags from a Jinja template.
        
        Args:
            template_filename: Name of template file
            
        Returns:
            Set of variable names found in template
        """
        try:
            with open(f"{self.template_dir}/{template_filename}", 'r', encoding='utf-8') as f:
                content = f.read()
            
            # Match {{ variable_name }} patterns
            pattern = r'\{\{\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*\}\}'
            variables = set(re.findall(pattern, content))
            
            logger.debug(f"Found {len(variables)} variables in template: {sorted(variables)}")
            return variables
        except Exception as e:
            logger.error(f"Error extracting template variables: {e}")
            return set()
    
    def validate_template_variables(
        self, 
        template_filename: str, 
        provided_vars: Dict[str, any],
        critical_vars: Set[str] = None
    ) -> tuple[list, list]:
        """Validate that all template variables are provided.
        
        Args:
            template_filename: Name of template file
            provided_vars: Dictionary of variables to be passed to template
            critical_vars: Set of variable names that are critical (missing = error)
            
        Returns:
            Tuple of (errors, warnings) - lists of missing variable names
        """
        template_vars = self.extract_template_variables(template_filename)
        provided_set = set(provided_vars.keys())
        
        missing = template_vars - provided_set
        
        errors = []
        warnings = []
        
        if critical_vars:
            critical_missing = missing & critical_vars
            if critical_missing:
                errors.extend(sorted(critical_missing))
                logger.error(f"Missing critical variables: {sorted(critical_missing)}")
            
            non_critical_missing = missing - critical_vars
            if non_critical_missing:
                warnings.extend(sorted(non_critical_missing))
                logger.warning(f"Missing non-critical variables: {sorted(non_critical_missing)}")
        else:
            if missing:
                warnings.extend(sorted(missing))
                logger.warning(f"Missing variables: {sorted(missing)}")
        
        return errors, warnings
    
    def render(
        self, 
        template_filename: str, 
        data: Dict[str, any], 
        output_path: str
    ) -> bool:
        """Render template with data and save to file.
        
        Args:
            template_filename: Name of template file
            data: Dictionary of variables to pass to template
            output_path: Path to save rendered output
            
        Returns:
            True if successful, False otherwise
        """
        try:
            template = self.load_template(template_filename)
            
            # Render template
            logger.info("Rendering template...")
            rendered = template.render(**data)
            
            # Write to file
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(rendered)
            
            logger.info(f"Ceremony script saved to: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error rendering template: {e}")
            return False
