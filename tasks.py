import os
import tempfile
import shutil
import zipfile
from werkzeug.utils import secure_filename
from worksheet_generator import generate_worksheets
from caterpillar_generator import generate_caterpillar_worksheets
from celery_app import celery_app

# Download folder configuration
DOWNLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'downloads')
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def zip_directory(folder_path, zip_path):
    """Helper to zip a directory"""
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname)

@celery_app.task(bind=True)
def generate_worksheets_task(self, file_path, generator_type, api_key):
    """
    Background task to generate worksheets
    
    Args:
        self: Celery task instance (bound)
        file_path: Path to uploaded Excel file
        generator_type: 'academy' or 'caterpillar'
        api_key: OpenAI API key
        
    Returns:
        dict with download URLs and generated files
    """
    try:
        # Create temporary directory for generation
        session_dir = tempfile.mkdtemp()
        output_dir = os.path.join(session_dir, 'output')
        os.makedirs(output_dir, exist_ok=True)
        
        # Select appropriate generator
        generator_func = generate_worksheets if generator_type == 'academy' else generate_caterpillar_worksheets
        
        # Track progress
        individual_files = []
        total_topics = 0
        completed_topics = 0
        
        # Update initial state
        self.update_state(
            state='PROGRESS',
            meta={
                'current': 0,
                'total': 100,
                'status': 'Starting generation...',
                'individual_files': []
            }
        )
        
        # Run generator and collect results
        for update in generator_func(file_path, output_dir, api_key):
            if update['type'] == 'progress':
                # Update progress state
                self.update_state(
                    state='PROGRESS',
                    meta={
                        'current': completed_topics,
                        'total': total_topics if total_topics > 0 else 100,
                        'status': update.get('message', 'Processing...'),
                        'individual_files': individual_files
                    }
                )
                
            elif update['type'] == 'result':
                # Zip individual topic
                topic_path = update['path']
                topic_name = update['topic']
                safe_topic_name = secure_filename(topic_name)
                zip_filename = f"{safe_topic_name}_{self.request.id}.zip"
                zip_path = os.path.join(DOWNLOAD_FOLDER, zip_filename)
                zip_directory(topic_path, zip_path)
                
                completed_topics += 1
                individual_files.append({
                    'topic': topic_name,
                    'filename': zip_filename,
                    'download_url': f'/download/{zip_filename}'
                })
                
                # Update progress
                progress_parts = update.get('progress', '0/0').split('/')
                if len(progress_parts) == 2:
                    completed_topics = int(progress_parts[0])
                    total_topics = int(progress_parts[1])
                
                self.update_state(
                    state='PROGRESS',
                    meta={
                        'current': completed_topics,
                        'total': total_topics,
                        'status': f'Completed: {topic_name}',
                        'individual_files': individual_files
                    }
                )
                
            elif update['type'] == 'complete':
                # Zip full output
                full_output_path = update['path']
                zip_filename = f"all_worksheets_{self.request.id}.zip"
                zip_path = os.path.join(DOWNLOAD_FOLDER, zip_filename)
                zip_directory(full_output_path, zip_path)
                
                # Clean up session directory
                shutil.rmtree(session_dir)
                
                # Return final result
                return {
                    'status': 'complete',
                    'download_url': f'/download/{zip_filename}',
                    'filename': zip_filename,
                    'individual_files': individual_files,
                    'total_generated': len(individual_files)
                }
        
        # If no complete signal received, still return what we have
        return {
            'status': 'complete',
            'individual_files': individual_files,
            'total_generated': len(individual_files)
        }
        
    except Exception as e:
        # Clean up on error
        if 'session_dir' in locals():
            shutil.rmtree(session_dir, ignore_errors=True)
        
        # Update state to failure
        self.update_state(
            state='FAILURE',
            meta={
                'status': 'Generation failed',
                'error': str(e)
            }
        )
        raise
