import React, { useRef, useState, useEffect } from 'react';
import './PasteArea.css';

interface PasteAreaProps {
    imagetabledata: (data: string) => void;
    containerClassName?: string;
}

const PasteArea: React.FC<PasteAreaProps> = ({ imagetabledata, containerClassName = '' }) => {
    const editorRef = useRef<HTMLDivElement | null>(null);
    const [content, setContent] = useState<string>('');

    // Sync initial content
    useEffect(() => {
        if (editorRef.current) {
            editorRef.current.innerHTML = content;
        }
    }, []);
    const handlePaste = (event: React.ClipboardEvent<HTMLDivElement>): void => {
        event.preventDefault();
        const items = Array.from(event.clipboardData.items);

        let htmlPromise: Promise<string | null> = Promise.resolve(null);
        let textPromise: Promise<string | null> = Promise.resolve(null);

        // Find HTML and plain text types
        items.forEach(item => {
            if (item.type === 'text/html') {
                htmlPromise = new Promise(resolve => {
                    item.getAsString((html: string) => resolve(html));
                });
            }
            if (item.type === 'text/plain') {
                textPromise = new Promise(resolve => {
                    item.getAsString((text: string) => resolve(text));
                });
            }
            if (item.type.indexOf('image') !== -1) {
                const blob = item.getAsFile();
                if (blob) {
                    const reader = new FileReader();
                    reader.onload = (e: ProgressEvent<FileReader>) => {
                        const img = document.createElement('img');
                        if (e.target?.result) {
                            img.src = e.target.result as string;
                            img.style.maxWidth = '300px';
                            const selection = window.getSelection();
                            if (selection && selection.rangeCount) {
                                const range = selection.getRangeAt(0);
                                range.insertNode(img);
                                range.collapse(false);
                            } else if (editorRef.current) {
                                editorRef.current.appendChild(img);
                            }
                            if (editorRef.current) {
                                setContent(editorRef.current.innerHTML);
                                imagetabledata(editorRef.current.innerHTML);
                            }
                        }
                    };
                    reader.readAsDataURL(blob);
                }
            }
        });

        // Now wait for html/text content and handle in correct order
        void Promise.all([htmlPromise, textPromise]).then(([htmlContent, plainText]) => {
            if (htmlContent) {
                const parser = new DOMParser();
                const doc = parser.parseFromString(htmlContent, 'text/html');
                const tables = doc.getElementsByTagName('table');

                if (tables.length > 0) {
                    const table = tables[0].cloneNode(true);
                    const selection = window.getSelection();
                    if (selection && selection.rangeCount) {
                        const range = selection.getRangeAt(0);
                        range.insertNode(table);
                        range.collapse(false);
                    } else if (editorRef.current) {
                        editorRef.current.appendChild(table);
                    }

                    if (editorRef.current) {
                        setContent(editorRef.current.innerHTML);
                        imagetabledata(editorRef.current.innerHTML);
                    }

                    return; // Stop here — don't insert plain text
                }
            }

            // If no table, insert plain text
            if (plainText) {
                const paragraphs = plainText.split('\n').filter(line => line.trim() !== '');
                const selection = window.getSelection();
                if (selection && selection.rangeCount) {
                    const range = selection.getRangeAt(0);
                    paragraphs.forEach((para, index) => {
                        const p = document.createElement('p');
                        p.textContent = para;
                        range.insertNode(p);
                        if (index < paragraphs.length - 1) {
                            range.collapse(false);
                            const br = document.createElement('br');
                            range.insertNode(br);
                        }
                    });
                    range.collapse(false);
                } else if (editorRef.current) {
                    paragraphs.forEach((para, index) => {
                        const p = document.createElement('p');
                        p.textContent = para;
                        editorRef.current!.appendChild(p);
                        if (index < paragraphs.length - 1) {
                            const br = document.createElement('br');
                            editorRef.current!.appendChild(br);
                        }
                    });
                }

                if (editorRef.current) {
                    setContent(editorRef.current.innerHTML);
                    imagetabledata(editorRef.current.innerHTML);
                }
            }
        });
    };

    const handleInput = (event: React.FormEvent<HTMLDivElement>): void => {
        setContent(event.currentTarget.innerHTML);
        imagetabledata(event.currentTarget.innerHTML)
        console.log("content------------------------------", event.currentTarget.innerHTML);
    };

    // Handle key events for better text input
    const handleKeyDown = (event: React.KeyboardEvent<HTMLDivElement>): void => {
        if (event.key === 'Enter' && !event.shiftKey) {
            event.preventDefault();
            document.execCommand('insertHTML', false, '<div><br></div>');
        }
    };

    return (
       <div className={`paste-container ${containerClassName}`}>
            <div
                ref={editorRef}
                contentEditable={true}
                onPaste={handlePaste}
                onInput={handleInput}
                onKeyDown={handleKeyDown}
                className="editor-area"
                data-placeholder="Type or paste text, images, or tables here..."
            />
            {/* <button onClick={() => { imagetabledata(content); }}>post</button> */}
        </div>
    );
};

export default PasteArea;