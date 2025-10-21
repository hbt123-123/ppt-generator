document.addEventListener('DOMContentLoaded', function() {
    // 当前项目状态
    let currentProject = {
        title: "论文答辩",
        subtitle: "基于环境工程的论文研究",
        presenter: "张三",
        date: new Date().toLocaleDateString('zh-CN', { year: 'numeric', month: 'long', day: 'numeric' }),
        team: "计算机科学与技术学院",
        theme: "spring",
        useSubtitle: true,
        includeToc: true,
        fontSize: 24,
        lineSpacing: 1.15,
        autoNumbering: true,
        thankYouText: "感谢观看\n提问环节\n如有任何问题，欢迎提问",
        slides: []
    };

    // 初始化UI
    initUI();

    // 初始化UI组件
    function initUI() {
        // 设置初始值
        document.getElementById('title').value = currentProject.title;
        document.getElementById('subtitle').value = currentProject.subtitle;
        document.getElementById('presenter').value = currentProject.presenter;
        document.getElementById('date').value = currentProject.date;
        document.getElementById('team').value = currentProject.team;
        document.getElementById('use-subtitle').checked = currentProject.useSubtitle;
        document.getElementById('include-toc').checked = currentProject.includeToc;
        document.getElementById('font-size').value = currentProject.fontSize;
        document.getElementById('line-spacing').value = currentProject.lineSpacing;
        document.getElementById('auto-numbering').checked = currentProject.autoNumbering;
        document.getElementById('thank-you-text').value = currentProject.thankYouText;

        // 主题选择
        document.querySelectorAll('.theme-card').forEach(card => {
            card.addEventListener('click', function() {
                document.querySelectorAll('.theme-card').forEach(c => c.classList.remove('active'));
                this.classList.add('active');
                currentProject.theme = this.dataset.theme;
            });
        });

        // 默认选择主题
        document.querySelector(`.theme-card[data-theme="${currentProject.theme}"]`).classList.add('active');

        // 添加幻灯片
        document.getElementById('add-slide').addEventListener('click', showSlideTypeModal);

        // 生成PPT
        document.getElementById('generate-btn').addEventListener('click', () => generatePPT(false));

        // 预览整个PPT
        document.getElementById('preview-btn').addEventListener('click', () => generatePPT(true));

        // 关闭预览模态框
        document.querySelector('.close').addEventListener('click', () => {
            document.getElementById('preview-modal').classList.add('hidden');
        });

        document.getElementById('close-preview').addEventListener('click', () => {
            document.getElementById('preview-modal').classList.add('hidden');
        });
    }

    // 显示幻灯片类型选择模态框
    function showSlideTypeModal() {
        document.getElementById('slide-type-modal').classList.remove('hidden');
    }

    // 隐藏幻灯片类型选择模态框
    function hideSlideTypeModal() {
        document.getElementById('slide-type-modal').classList.add('hidden');
    }

    // 添加新幻灯片
    function addNewSlide(type = 'content', title = '') {
        const slideTemplate = document.getElementById('slide-template');
        const slideClone = slideTemplate.content.cloneNode(true);
        const slideElement = slideClone.querySelector('.slide-editor');

        // 设置唯一ID
        const slideId = `slide-${Date.now()}`;
        slideElement.dataset.slideId = slideId;

        // 设置幻灯片类型
        const typeSelect = slideElement.querySelector('.slide-type');
        typeSelect.value = type;

        // 设置幻灯片标题
        const titleInput = slideElement.querySelector('.slide-title');
        titleInput.value = title || getDefaultTitle(type);

        // 添加到容器
        const slidesContainer = document.getElementById('slides-container');

        // 移除初始占位符
        const placeholder = document.getElementById('slides-placeholder');
        if (placeholder) placeholder.remove();

        slidesContainer.appendChild(slideClone);

        // 创建幻灯片数据对象
        const slideData = {
            id: slideId,
            type: type,
            title: titleInput.value,
            content: [],
            left_content: [],
            right_content: [],
            includeInToc: true
        };

        // 添加到项目数据
        currentProject.slides.push(slideData);

        // 设置事件监听器
        setupSlideEvents(slideElement, slideData);

        // 根据类型调整UI
        adjustUIForType(type, slideElement);

        // 添加默认内容项
        addDefaultContent(type, slideElement, slideData);
    }

    // 创建内容项
    function createContentItem(content = '', containerClass = '') {
        const itemTemplate = document.getElementById('content-item-template');
        const itemClone = itemTemplate.content.cloneNode(true);
        const contentInput = itemClone.querySelector('.content-input');
        contentInput.value = content;

        // 添加删除事件
        const deleteBtn = itemClone.querySelector('.btn-delete');
        deleteBtn.addEventListener('click', function() {
            const item = this.closest('.content-item');
            item.remove();

            // 更新项目数据
            const slideElement = item.closest('.slide-editor');
            const slideId = slideElement.dataset.slideId;
            const slideData = currentProject.slides.find(s => s.id === slideId);

            if (slideData) {
                if (containerClass === 'left-items') {
                    slideData.left_content = Array.from(slideElement.querySelectorAll('.left-items .content-input'))
                        .map(input => input.value);
                } else if (containerClass === 'right-items') {
                    slideData.right_content = Array.from(slideElement.querySelectorAll('.right-items .content-input'))
                        .map(input => input.value);
                } else {
                    slideData.content = Array.from(slideElement.querySelectorAll('.content-items .content-input'))
                        .map(input => input.value);
                }
            }
        });

        return itemClone;
    }

    // 设置幻灯片事件
    function setupSlideEvents(slideElement, slideData) {
        // 幻灯片类型切换
        const typeSelect = slideElement.querySelector('.slide-type');
        typeSelect.addEventListener('change', function() {
            slideData.type = this.value;
            adjustUIForType(this.value, slideElement);
        });

        // 标题更新
        const titleInput = slideElement.querySelector('.slide-title');
        titleInput.addEventListener('input', function() {
            slideData.title = this.value;
        });

        // 副标题更新
        const subtitleInput = slideElement.querySelector('.slide-subtitle');
        if (subtitleInput) {
            subtitleInput.addEventListener('input', function() {
                slideData.subtitle = this.value;
            });
        }

        // 左侧副标题更新
        const leftSubtitleInput = slideElement.querySelector('.left-subtitle');
        if (leftSubtitleInput) {
            leftSubtitleInput.addEventListener('input', function() {
                slideData.left_subtitle = this.value;
            });
        }

        // 右侧副标题更新
        const rightSubtitleInput = slideElement.querySelector('.right-subtitle');
        if (rightSubtitleInput) {
            rightSubtitleInput.addEventListener('input', function() {
                slideData.right_subtitle = this.value;
            });
        }

        // 视频文本更新
        const videoTextInput = slideElement.querySelector('.video-text');
        if (videoTextInput) {
            videoTextInput.addEventListener('input', function() {
                slideData.video_text = this.value;
            });
        }

        // 添加内容项
        const addItemBtn = slideElement.querySelector('.add-content-item');
        if (addItemBtn) {
            addItemBtn.addEventListener('click', function() {
                const contentItems = slideElement.querySelector('.content-items');
                const newItem = createContentItem('', 'content-items');
                contentItems.appendChild(newItem);
                slideData.content.push('');
            });
        }

        // 添加左侧内容项
        const addLeftItemBtn = slideElement.querySelector('.add-left-item');
        if (addLeftItemBtn) {
            addLeftItemBtn.addEventListener('click', function() {
                const leftItems = slideElement.querySelector('.left-items');
                const newItem = createContentItem('', 'left-items');
                leftItems.appendChild(newItem);
                slideData.left_content.push('');
            });
        }

        // 添加右侧内容项
        const addRightItemBtn = slideElement.querySelector('.add-right-item');
        if (addRightItemBtn) {
            addRightItemBtn.addEventListener('click', function() {
                const rightItems = slideElement.querySelector('.right-items');
                const newItem = createContentItem('', 'right-items');
                rightItems.appendChild(newItem);
                slideData.right_content.push('');
            });
        }

        // 图片上传
        const imageInput = slideElement.querySelector('.image-input');
        if (imageInput) {
            imageInput.addEventListener('change', function(e) {
                if (e.target.files && e.target.files[0]) {
                    const reader = new FileReader();
                    reader.onload = function(event) {
                        const imagePreview = slideElement.querySelector('.image-preview');
                        imagePreview.innerHTML = `<img src="${event.target.result}" alt="预览">`;
                        imagePreview.classList.remove('hidden');
                        slideData.image = event.target.result;
                    };
                    reader.readAsDataURL(e.target.files[0]);
                }
            });
        }

        // 左侧图片上传
        const leftImageInput = slideElement.querySelector('.left-image-input');
        if (leftImageInput) {
            leftImageInput.addEventListener('change', function(e) {
                if (e.target.files && e.target.files[0]) {
                    const reader = new FileReader();
                    reader.onload = function(event) {
                        const imagePreview = slideElement.querySelector('.left-image-preview');
                        imagePreview.innerHTML = `<img src="${event.target.result}" alt="预览">`;
                        imagePreview.classList.remove('hidden');
                        slideData.left_image = event.target.result;
                    };
                    reader.readAsDataURL(e.target.files[0]);
                }
            });
        }

        // 右侧图片上传
        const rightImageInput = slideElement.querySelector('.right-image-input');
        if (rightImageInput) {
            rightImageInput.addEventListener('change', function(e) {
                if (e.target.files && e.target.files[0]) {
                    const reader = new FileReader();
                    reader.onload = function(event) {
                        const imagePreview = slideElement.querySelector('.right-image-preview');
                        imagePreview.innerHTML = `<img src="${event.target.result}" alt="预览">`;
                        imagePreview.classList.remove('hidden');
                        slideData.right_image = event.target.result;
                    };
                    reader.readAsDataURL(e.target.files[0]);
                }
            });
        }

        // 视频URL输入
        const videoUrlInput = slideElement.querySelector('.video-url');
        if (videoUrlInput) {
            videoUrlInput.addEventListener('input', function() {
                slideData.video_url = this.value;
            });
        }

        // 删除按钮
        const deleteBtn = slideElement.querySelector('.btn-delete-slide');
        deleteBtn.addEventListener('click', function() {
            if (confirm('确定要删除这张幻灯片吗？')) {
                slideElement.remove();
                currentProject.slides = currentProject.slides.filter(s => s.id !== slideData.id);
                if (!document.querySelectorAll('.slide-editor').length) showSlidesPlaceholder();
            }
        });

        // 上移按钮
        const moveUpBtn = slideElement.querySelector('.btn-move-up');
        moveUpBtn.addEventListener('click', function() {
            const container = document.getElementById('slides-container');
            const slides = Array.from(container.querySelectorAll('.slide-editor'));
            const index = slides.indexOf(slideElement);

            if (index > 0) {
                container.insertBefore(slideElement, slides[index - 1]);

                // 更新数据顺序
                const slideIndex = currentProject.slides.findIndex(s => s.id === slideData.id);
                if (slideIndex > 0) {
                    [currentProject.slides[slideIndex], currentProject.slides[slideIndex - 1]] =
                    [currentProject.slides[slideIndex - 1], currentProject.slides[slideIndex]];
                }
            }
        });

        // 下移按钮
        const moveDownBtn = slideElement.querySelector('.btn-move-down');
        moveDownBtn.addEventListener('click', function() {
            const container = document.getElementById('slides-container');
            const slides = Array.from(container.querySelectorAll('.slide-editor'));
            const index = slides.indexOf(slideElement);

            if (index < slides.length - 1) {
                container.insertBefore(slides[index + 1], slideElement);

                // 更新数据顺序
                const slideIndex = currentProject.slides.findIndex(s => s.id === slideData.id);
                if (slideIndex < currentProject.slides.length - 1) {
                    [currentProject.slides[slideIndex], currentProject.slides[slideIndex + 1]] =
                    [currentProject.slides[slideIndex + 1], currentProject.slides[slideIndex]];
                }
            }
        });

        // 包含在目录中
        const includeInToc = slideElement.querySelector('.include-in-toc');
        includeInToc.addEventListener('change', function() {
            slideData.includeInToc = this.checked;
        });

        // 内容输入更新
        slideElement.querySelectorAll('.content-input').forEach(input => {
            input.addEventListener('input', function() {
                const container = this.closest('.content-items, .left-items, .right-items');
                const slideId = slideElement.dataset.slideId;
                const slideData = currentProject.slides.find(s => s.id === slideId);

                if (slideData) {
                    if (container.classList.contains('left-items')) {
                        slideData.left_content = Array.from(container.querySelectorAll('.content-input'))
                            .map(input => input.value);
                    } else if (container.classList.contains('right-items')) {
                        slideData.right_content = Array.from(container.querySelectorAll('.content-input'))
                            .map(input => input.value);
                    } else {
                        slideData.content = Array.from(container.querySelectorAll('.content-input'))
                            .map(input => input.value);
                    }
                }
            });
        });
    }

    // 根据幻灯片类型调整UI
    function adjustUIForType(type, slideElement) {
        // 隐藏所有内容区域
        slideElement.querySelectorAll('.content-group, .media-group').forEach(group => {
            group.classList.add('hidden');
        });

        // 显示特定内容区域
        switch(type) {
            case 'title_subtitle':
                slideElement.querySelector('.subtitle-group').classList.remove('hidden');
                break;

            case 'section_header':
                slideElement.querySelector('.subtitle-group').classList.remove('hidden');
                break;

            case 'content':
                slideElement.querySelector('.single-content-group').classList.remove('hidden');
                break;

            case 'two_content':
                slideElement.querySelector('.two-content-group').classList.remove('hidden');
                break;

            case 'comparison':
                slideElement.querySelector('.comparison-group').classList.remove('hidden');
                break;

            case 'content_caption':
                slideElement.querySelector('.single-content-group').classList.remove('hidden');
                slideElement.querySelector('.image-upload-group').classList.remove('hidden');
                break;

            case 'picture_caption':
                slideElement.querySelector('.image-upload-group').classList.remove('hidden');
                break;

            case 'video':
                slideElement.querySelector('.video-upload-group').classList.remove('hidden');
                break;
        }
    }

    // 添加默认内容
    function addDefaultContent(type, slideElement, slideData) {
        const defaults = {
            'content': ['研究背景', '研究方法', '研究结果'],
            'two_content': ['左侧要点1', '左侧要点2'],
            'comparison': ['左侧对比点1', '左侧对比点2'],
            'content_caption': ['内容要点1', '内容要点2']
        };

        if (defaults[type]) {
            if (type === 'two_content' || type === 'comparison') {
                // 添加左右内容
                defaults[type].forEach((content, i) => {
                    if (i % 2 === 0) {
                        const leftItems = slideElement.querySelector('.left-items');
                        const newItem = createContentItem(content, 'left-items');
                        leftItems.appendChild(newItem);
                        slideData.left_content.push(content);
                    } else {
                        const rightItems = slideElement.querySelector('.right-items');
                        const newItem = createContentItem(content, 'right-items');
                        rightItems.appendChild(newItem);
                        slideData.right_content.push(content);
                    }
                });
            } else {
                // 添加单列内容
                defaults[type].forEach(content => {
                    const contentItems = slideElement.querySelector('.content-items');
                    const newItem = createContentItem(content, 'content-items');
                    contentItems.appendChild(newItem);
                    slideData.content.push(content);
                });
            }
        }
    }

    // 获取默认标题
    function getDefaultTitle(type) {
        const titles = {
            'title_subtitle': '标题幻灯片',
            'section_header': '研究背景与意义',
            'content': '主要内容',
            'two_content': '两栏内容',
            'comparison': '比较分析',
            'content_caption': '内容与标题',
            'picture_caption': '图片与标题',
            'video': '视频演示'
        };
        return titles[type] || '新幻灯片';
    }

    // 显示幻灯片占位符
    function showSlidesPlaceholder() {
        const slidesContainer = document.getElementById('slides-container');
        const placeholder = document.createElement('div');
        placeholder.className = 'placeholder';
        placeholder.id = 'slides-placeholder';
        placeholder.innerHTML = `
            <i class="fas fa-file-powerpoint"></i>
            <h3>开始创建您的演示文稿</h3>
            <p>点击"添加幻灯片"按钮开始创建您的第一张幻灯片</p>
            <p>您可以在右侧设置封面信息和主题样式</p>
        `;
        slidesContainer.appendChild(placeholder);
    }

    // 收集项目数据
    function collectProjectData() {
        currentProject.title = document.getElementById('title').value;
        currentProject.subtitle = document.getElementById('subtitle').value;
        currentProject.presenter = document.getElementById('presenter').value;
        currentProject.date = document.getElementById('date').value;
        currentProject.team = document.getElementById('team').value;
        currentProject.useSubtitle = document.getElementById('use-subtitle').checked;
        currentProject.includeToc = document.getElementById('include-toc').checked;
        currentProject.fontSize = parseInt(document.getElementById('font-size').value);
        currentProject.lineSpacing = parseFloat(document.getElementById('line-spacing').value);
        currentProject.autoNumbering = document.getElementById('auto-numbering').checked;
        currentProject.thankYouText = document.getElementById('thank-you-text').value;
    }

    // 生成PPT
    function generatePPT(isPreview = false) {
        collectProjectData();

        // 验证必填项
        if (!currentProject.title.trim()) {
            alert('请填写演示文稿标题');
            return;
        }

        if (!currentProject.presenter.trim()) {
            alert('请填写主讲人姓名');
            return;
        }

        // 显示加载状态
        const generateBtn = document.getElementById('generate-btn');
        const originalBtnText = generateBtn.innerHTML;
        generateBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> 生成中...';
        generateBtn.disabled = true;

        if (isPreview) {
            document.getElementById('preview-modal').classList.remove('hidden');
            document.getElementById('modal-placeholder').classList.remove('hidden');
            document.getElementById('ppt-preview').classList.add('hidden');
        }

        // 发送请求
        fetch('/generate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(currentProject),
        })
        .then(response => {
            if (!response.ok) throw new Error('网络响应错误');
            return response.blob();
        })
        .then(blob => {
            if (isPreview) {
                const url = URL.createObjectURL(blob);
                const previewFrame = document.getElementById('ppt-preview');
                previewFrame.src = url;

                document.getElementById('modal-placeholder').classList.add('hidden');
                previewFrame.classList.remove('hidden');

                document.getElementById('download-preview').onclick = function() {
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = currentProject.title.replace(/ /g, '_') + '.pptx';
                    document.body.appendChild(a);
                    a.click();
                    a.remove();
                };
            } else {
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = currentProject.title.replace(/ /g, '_') + '.pptx';
                document.body.appendChild(a);
                a.click();
                a.remove();
            }

            generateBtn.innerHTML = originalBtnText;
            generateBtn.disabled = false;
        })
        .catch(error => {
            console.error('生成PPT时出错:', error);
            alert('生成PPT失败: ' + error.message);
            generateBtn.innerHTML = originalBtnText;
            generateBtn.disabled = false;

            if (isPreview) {
                document.getElementById('modal-placeholder').innerHTML = `
                    <i class="fas fa-exclamation-triangle"></i>
                    <h3>生成失败</h3>
                    <p>${error.message}</p>
                    <p>请检查网络连接后重试</p>
                `;
            }
        });
    }

    // 初始化幻灯片类型选择事件
    document.querySelectorAll('.slide-type-card').forEach(card => {
        card.addEventListener('click', function() {
            const type = this.dataset.type;
            hideSlideTypeModal();
            addNewSlide(type);
        });
    });

    // 关闭幻灯片类型选择模态框
    document.querySelector('.close-type-modal').addEventListener('click', hideSlideTypeModal);
});