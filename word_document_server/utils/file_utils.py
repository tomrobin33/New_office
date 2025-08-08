"""
File utility functions for Word Document Server.
"""
import os
from typing import Tuple, Optional
import shutil
import requests
import paramiko


def check_file_writeable(filepath: str) -> Tuple[bool, str]:
    """
    Check if a file can be written to.
    
    Args:
        filepath: Path to the file
        
    Returns:
        Tuple of (is_writeable, error_message)
    """
    # If file doesn't exist, check if directory is writeable
    if not os.path.exists(filepath):
        directory = os.path.dirname(filepath)
        # If no directory is specified (empty string), use current directory
        if directory == '':
            directory = '.'
        if not os.path.exists(directory):
            return False, f"Directory {directory} does not exist"
        if not os.access(directory, os.W_OK):
            return False, f"Directory {directory} is not writeable"
        return True, ""
    
    # If file exists, check if it's writeable
    if not os.access(filepath, os.W_OK):
        return False, f"File {filepath} is not writeable (permission denied)"
    
    # Try to open the file for writing to see if it's locked
    try:
        with open(filepath, 'a'):
            pass
        return True, ""
    except IOError as e:
        return False, f"File {filepath} is not writeable: {str(e)}"
    except Exception as e:
        return False, f"Unknown error checking file permissions: {str(e)}"


def create_document_copy(source_path: str, dest_path: Optional[str] = None) -> Tuple[bool, str, Optional[str]]:
    """
    Create a copy of a document.
    
    Args:
        source_path: Path to the source document
        dest_path: Optional path for the new document. If not provided, will use source_path + '_copy.docx'
        
    Returns:
        Tuple of (success, message, new_filepath)
    """
    if not os.path.exists(source_path):
        return False, f"Source document {source_path} does not exist", None
    
    if not dest_path:
        # Generate a new filename if not provided
        base, ext = os.path.splitext(source_path)
        dest_path = f"{base}_copy{ext}"
    
    try:
        # Simple file copy
        shutil.copy2(source_path, dest_path)
        return True, f"Document copied to {dest_path}", dest_path
    except Exception as e:
        return False, f"Failed to copy document: {str(e)}", None


def ensure_docx_extension(filename: str) -> str:
    """
    Ensure filename has .docx extension.
    
    Args:
        filename: The filename to check
        
    Returns:
        Filename with .docx extension
    """
    if not filename.endswith('.docx'):
        return filename + '.docx'
    return filename


def download_file_from_url(url: str, save_dir: str = ".") -> str:
    """
    从URL下载文件到本地指定目录，返回本地文件路径。
    """
    if not os.path.exists(save_dir):
        os.makedirs(save_dir, exist_ok=True)
    
    # 清理URL，移除末尾的斜杠
    cleaned_url = url.rstrip('/')
    
    # 从URL中提取文件名
    filename = cleaned_url.split("/")[-1].split("?")[0]
    
    # 如果文件名为空，生成一个默认文件名
    if not filename:
        import hashlib
        filename = f"downloaded_file_{hashlib.md5(cleaned_url.encode()).hexdigest()[:8]}"
    
    local_filename = os.path.join(save_dir, filename)
    
    # 设置请求头，模拟浏览器请求
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'application/pdf,application/octet-stream,*/*',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
        'Accept-Encoding': 'gzip, deflate',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
    }
    
    try:
        # 使用更宽松的请求设置
        r = requests.get(cleaned_url, stream=True, headers=headers, timeout=30, allow_redirects=True)
        r.raise_for_status()
        
        with open(local_filename, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
        
        return local_filename
        
    except requests.exceptions.RequestException as e:
        # 如果第一次请求失败，尝试使用原始URL（不清理斜杠）
        try:
            r = requests.get(url, stream=True, headers=headers, timeout=30, allow_redirects=True)
            r.raise_for_status()
            
            with open(local_filename, 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            return local_filename
            
        except requests.exceptions.RequestException as e2:
            raise Exception(f"下载文件失败: {str(e2)}")


def upload_file_to_server(local_path: str, remote_path: str, server: str, username: str, password: str) -> str:
    """
    通过SFTP上传文件到远程服务器。
    """
    transport = paramiko.Transport((server, 22))
    transport.connect(username=username, password=password)
    sftp = paramiko.SFTPClient.from_transport(transport)
    sftp.put(local_path, remote_path)
    sftp.close()
    transport.close()
    return f"上传成功: {remote_path}"
