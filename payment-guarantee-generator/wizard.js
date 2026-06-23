/**
 * Wizard.js - 通用向导式翻页框架
 * 用于合同自动生成网站，管理多步骤表单的导航、进度和 PDF 生成。
 *
 * HTML 结构约定：
 *   .wizard-progress       - 进度条容器
 *     .wizard-progress-bar - 进度条填充
 *     .wizard-progress-text- 进度文字
 *   .wizard-indicator      - 步骤指示器容器
 *   .wizard-content        - 步骤内容容器
 *     .wizard-step[data-wizard-step] - 各步骤面板
 *   .wizard-nav            - 导航按钮容器
 *     .wizard-btn-prev     - 上一步按钮
 *     .wizard-btn-next     - 下一步按钮
 *     .wizard-btn-generate - 生成 PDF 按钮
 *
 * 自定义事件：
 *   wizard:stepChange  - 步骤切换时触发，detail: { currentStep, totalSteps, previousStep }
 *   wizard:beforeNext  - 点击"下一步"前触发，可 preventDefault() 阻止跳转
 *   wizard:beforePrev  - 点击"上一步"前触发，可 preventDefault() 阻止跳转
 *   wizard:generate    - 开始生成 PDF 时触发
 *   wizard:generated   - PDF 生成完成时触发，detail: { filename }
 *   wizard:generateError - PDF 生成失败时触发，detail: { error }
 */
;(function (global) {
  'use strict';

  // ──────────────────────────── 默认配置 ────────────────────────────
  var DEFAULTS = {
    formId: 'applicationForm',
    generateUrl: '/generate',
    generateCallback: null,
    steps: null,           // 手动步骤名称数组，null 则从 DOM 读取
    autoInit: true,         // 是否在 DOMContentLoaded 时自动初始化
    keyboardNav: true,       // 是否启用键盘快捷键
    scrollRestore: true,     // 是否保存/恢复各步骤滚动位置
    animationDuration: 300   // 步骤切换动画时长 (ms)
  };

  // ──────────────────────────── 内部状态 ────────────────────────────
  var _options = {};
  var _currentStep = 0;
  var _totalSteps = 0;
  var _stepNames = [];
  var _scrollPositions = {};
  var _initialized = false;
  var _destroyed = false;

  // DOM 缓存
  var _els = {
    form: null,
    progress: null,
    progressBar: null,
    progressText: null,
    indicator: null,
    content: null,
    nav: null,
    btnPrev: null,
    btnNext: null,
    btnGenerate: null,
    steps: []
  };

  // ──────────────────────────── 工具函数 ────────────────────────────

  /**
   * 简单的类名操作
   */
  function addClass(el, cls) {
    if (el && el.classList) { el.classList.add(cls); }
  }

  function removeClass(el, cls) {
    if (el && el.classList) { el.classList.remove(cls); }
  }

  function hasClass(el, cls) {
    return el && el.classList ? el.classList.contains(cls) : false;
  }

  /**
   * 触发自定义事件
   */
  function emit(eventName, detail) {
    var event = new CustomEvent(eventName, {
      bubbles: true,
      cancelable: eventName.indexOf('before') === 0,
      detail: detail || {}
    });
    document.dispatchEvent(event);
    return event;
  }

  /**
   * 收集表单数据为 FormData
   */
  function collectFormData() {
    var form = _els.form;
    if (!form) {
      console.warn('[Wizard] 未找到表单元素，无法收集数据。');
      return new FormData();
    }
    return new FormData(form);
  }

  /**
   * 显示加载状态
   */
  function showLoading() {
    if (_els.btnGenerate) {
      _els.btnGenerate.disabled = true;
      _els.btnGenerate.dataset.originalText = _els.btnGenerate.textContent;
      _els.btnGenerate.textContent = '生成中...';
    }
    if (_els.btnNext) { _els.btnNext.disabled = true; }
    if (_els.btnPrev) { _els.btnPrev.disabled = true; }
  }

  /**
   * 隐藏加载状态
   */
  function hideLoading() {
    if (_els.btnGenerate) {
      _els.btnGenerate.disabled = false;
      if (_els.btnGenerate.dataset.originalText) {
        _els.btnGenerate.textContent = _els.btnGenerate.dataset.originalText;
      }
    }
    if (_els.btnNext) { _els.btnNext.disabled = false; }
    if (_els.btnPrev) { _els.btnPrev.disabled = false; }
  }

  // ──────────────────────────── DOM 查询与缓存 ────────────────────────────

  function cacheDOMElements() {
    _els.form = document.getElementById(_options.formId) || document.querySelector('form');
    _els.progress = document.querySelector('.wizard-progress');
    _els.progressBar = document.querySelector('.wizard-progress-bar');
    _els.progressText = document.querySelector('.wizard-progress-text');
    _els.indicator = document.querySelector('.wizard-indicator');
    _els.content = document.querySelector('.wizard-content');
    _els.nav = document.querySelector('.wizard-nav');
    _els.btnPrev = document.querySelector('.wizard-btn-prev');
    _els.btnNext = document.querySelector('.wizard-btn-next');
    _els.btnGenerate = document.querySelector('.wizard-btn-generate');
  }

  function collectSteps() {
    var stepElements = document.querySelectorAll('.wizard-step');
    if (stepElements.length === 0) { return; }

    // 转为数组并按 data-wizard-step 排序
    var arr = Array.prototype.slice.call(stepElements);
    arr.sort(function (a, b) {
      var aVal = parseInt(a.getAttribute('data-wizard-step'), 10) || 0;
      var bVal = parseInt(b.getAttribute('data-wizard-step'), 10) || 0;
      return aVal - bVal;
    });

    _els.steps = arr;
    _totalSteps = arr.length;

    // 收集步骤名称
    if (_options.steps && _options.steps.length === _totalSteps) {
      _stepNames = _options.steps.slice();
    } else {
      _stepNames = arr.map(function (el, i) {
        return el.getAttribute('data-wizard-title') || el.getAttribute('title') || ('步骤 ' + (i + 1));
      });
    }
  }

  // ──────────────────────────── 指示器渲染 ────────────────────────────

  function renderIndicator() {
    if (!_els.indicator) { return; }

    var html = '';
    for (var i = 0; i < _totalSteps; i++) {
      var stateClass = i === 0 ? ' active' : '';
      html += '<div class="wizard-indicator-item' + stateClass + '" data-indicator-index="' + i + '">';
      html += '  <div class="wizard-indicator-number">' + (i + 1) + '</div>';
      html += '  <div class="wizard-indicator-label">' + _stepNames[i] + '</div>';
      html += '</div>';
      if (i < _totalSteps - 1) {
        html += '<div class="wizard-indicator-line"></div>';
      }
    }
    _els.indicator.innerHTML = html;
  }

  // ──────────────────────────── 核心：步骤切换 ────────────────────────────

  function showStep(index, previousIndex) {
    // 保存当前步骤的滚动位置
    if (_options.scrollRestore && previousIndex !== undefined) {
      _scrollPositions[previousIndex] = window.pageYOffset || document.documentElement.scrollTop;
    }

    // 隐藏所有步骤
    _els.steps.forEach(function (el, i) {
      if (i === index) {
        removeClass(el, 'wizard-step-hidden');
        addClass(el, 'wizard-step-active');
      } else {
        removeClass(el, 'wizard-step-active');
        addClass(el, 'wizard-step-hidden');
      }
    });

    // 更新进度
    updateProgress();

    // 更新指示器
    updateIndicator();

    // 更新导航
    updateNav();

    // 恢复滚动位置
    if (_options.scrollRestore && _scrollPositions[index] !== undefined) {
      window.scrollTo(0, _scrollPositions[index]);
    } else {
      window.scrollTo(0, 0);
    }

    // 触发自定义事件
    emit('wizard:stepChange', {
      currentStep: index,
      totalSteps: _totalSteps,
      previousStep: previousIndex !== undefined ? previousIndex : -1
    });
  }

  // ──────────────────────────── 公共 API 方法 ────────────────────────────

  /**
   * 初始化向导
   * @param {Object} options - 配置项
   */
  function init(options) {
    if (_initialized) {
      console.warn('[Wizard] 已经初始化，请先调用 destroy() 再重新初始化。');
      return;
    }

    // 合并配置
    _options = {};
    for (var key in DEFAULTS) {
      if (DEFAULTS.hasOwnProperty(key)) {
        _options[key] = DEFAULTS[key];
      }
    }
    if (options && typeof options === 'object') {
      for (var optKey in options) {
        if (options.hasOwnProperty(optKey)) {
          _options[optKey] = options[optKey];
        }
      }
    }

    // 缓存 DOM
    cacheDOMElements();

    // 收集步骤
    collectSteps();
    if (_totalSteps === 0) {
      console.info('[Wizard] 未找到 .wizard-step 元素，跳过初始化（兼容旧页面）。');
      return;
    }

    _initialized = true;
    _destroyed = false;
    _currentStep = 0;
    _scrollPositions = {};

    // 渲染指示器
    renderIndicator();

    // 绑定键盘快捷键
    if (_options.keyboardNav) {
      document.addEventListener('keydown', _onKeyDown);
    }

    // 绑定指示器点击事件（事件委托）
    if (_els.indicator) {
      _els.indicator.addEventListener('click', _onIndicatorClick);
    }

    // 显示第一步
    showStep(0);

    console.log('[Wizard] 初始化完成，共 ' + _totalSteps + ' 个步骤。');
  }

  /**
   * 下一步
   */
  function nextStep() {
    if (!_initialized || _destroyed) { return; }

    // 触发 beforeNext 事件，允许拦截
    var beforeEvent = emit('wizard:beforeNext', {
      currentStep: _currentStep,
      totalSteps: _totalSteps
    });
    if (beforeEvent.defaultPrevented) { return; }

    if (_currentStep < _totalSteps - 1) {
      var prev = _currentStep;
      _currentStep++;
      showStep(_currentStep, prev);
    }
  }

  /**
   * 上一步
   */
  function prevStep() {
    if (!_initialized || _destroyed) { return; }

    // 触发 beforePrev 事件，允许拦截
    var beforeEvent = emit('wizard:beforePrev', {
      currentStep: _currentStep,
      totalSteps: _totalSteps
    });
    if (beforeEvent.defaultPrevented) { return; }

    if (_currentStep > 0) {
      var prev = _currentStep;
      _currentStep--;
      showStep(_currentStep, prev);
    }
  }

  /**
   * 跳转到指定步骤
   * @param {number} n - 目标步骤索引（从 0 开始）
   */
  function goToStep(n) {
    if (!_initialized || _destroyed) { return; }

    var target = parseInt(n, 10);
    if (isNaN(target) || target < 0 || target >= _totalSteps) {
      console.warn('[Wizard] goToStep: 无效的步骤索引 ' + n + '，有效范围 0-' + (_totalSteps - 1) + '。');
      return;
    }

    if (target === _currentStep) { return; }

    var prev = _currentStep;
    _currentStep = target;
    showStep(_currentStep, prev);
  }

  /**
   * 获取当前步骤索引（从 0 开始）
   */
  function getCurrentStep() {
    return _currentStep;
  }

  /**
   * 获取总步骤数
   */
  function getTotalSteps() {
    return _totalSteps;
  }

  /**
   * 更新进度条
   */
  function updateProgress() {
    if (_els.progressBar) {
      var percent = _totalSteps > 1 ? ((_currentStep + 1) / _totalSteps) * 100 : 100;
      _els.progressBar.style.width = percent + '%';
    }
    if (_els.progressText) {
      _els.progressText.textContent = '步骤 ' + (_currentStep + 1) + ' / ' + _totalSteps;
    }
  }

  /**
   * 更新步骤指示器状态
   */
  function updateIndicator() {
    if (!_els.indicator) { return; }

    var items = _els.indicator.querySelectorAll('.wizard-indicator-item');
    items.forEach(function (item, i) {
      removeClass(item, 'active');
      removeClass(item, 'completed');

      if (i === _currentStep) {
        addClass(item, 'active');
      } else if (i < _currentStep) {
        addClass(item, 'completed');
      }
    });

    // 更新连接线状态
    var lines = _els.indicator.querySelectorAll('.wizard-indicator-line');
    lines.forEach(function (line, i) {
      if (i < _currentStep) {
        addClass(line, 'completed');
      } else {
        removeClass(line, 'completed');
      }
    });
  }

  /**
   * 更新底部导航按钮
   */
  function updateNav() {
    // 第一步隐藏"上一步"按钮
    if (_els.btnPrev) {
      if (_currentStep === 0) {
        _els.btnPrev.style.display = 'none';
        _els.btnPrev.disabled = true;
      } else {
        _els.btnPrev.style.display = '';
        _els.btnPrev.disabled = false;
      }
    }

    // 最后一步隐藏"下一步"，显示"生成 PDF"
    if (_els.btnNext) {
      if (_currentStep === _totalSteps - 1) {
        _els.btnNext.style.display = 'none';
      } else {
        _els.btnNext.style.display = '';
      }
    }

    if (_els.btnGenerate) {
      if (_currentStep === _totalSteps - 1) {
        _els.btnGenerate.style.display = '';
      } else {
        _els.btnGenerate.style.display = 'none';
      }
    }
  }

  /**
   * 生成 PDF
   */
  function generate() {
    if (!_initialized || _destroyed) { return; }

    var formData = collectFormData();

    // 触发 generate 事件
    emit('wizard:generate', { formData: formData });

    showLoading();

    var params = new URLSearchParams();
    formData.forEach(function (value, key) {
      params.append(key, value);
    });

    fetch(_options.generateUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
      body: params.toString()
    })
    .then(function (response) {
      if (!response.ok) {
        return response.text().then(function (text) {
          throw new Error(text || '服务器返回错误: ' + response.status);
        });
      }

      // 尝试从 Content-Disposition 头提取文件名
      var disposition = response.headers.get('Content-Disposition') || '';
      var filenameMatch = disposition.match(/filename\*?=(?:UTF-8'')?([^;\s]+)/i);
      var filename = filenameMatch ? decodeURIComponent(filenameMatch[1].replace(/['"]/g, '')) : '合同.pdf';

      return response.blob().then(function (blob) {
        return { blob: blob, filename: filename };
      });
    })
    .then(function (result) {
      hideLoading();

      // 创建下载链接
      var url = URL.createObjectURL(result.blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = result.filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(function () { URL.revokeObjectURL(url); }, 1000);

      // 触发完成事件
      emit('wizard:generated', { filename: result.filename, blob: result.blob });

      // 调用自定义回调
      if (typeof _options.generateCallback === 'function') {
        _options.generateCallback(result.blob, result.filename);
      }
    })
    .catch(function (err) {
      hideLoading();
      console.error('[Wizard] PDF 生成失败:', err);

      emit('wizard:generateError', { error: err.message || err });

      alert('PDF 生成失败: ' + (err.message || '未知错误'));
    });
  }

  /**
   * 销毁向导实例，清理事件和状态
   */
  function destroy() {
    if (!_initialized) { return; }

    if (_options.keyboardNav) {
      document.removeEventListener('keydown', _onKeyDown);
    }

    if (_els.indicator) {
      _els.indicator.removeEventListener('click', _onIndicatorClick);
    }

    // 清除所有步骤的 active/hidden 类
    _els.steps.forEach(function (el) {
      removeClass(el, 'wizard-step-active');
      removeClass(el, 'wizard-step-hidden');
    });

    // 重置状态
    _currentStep = 0;
    _totalSteps = 0;
    _stepNames = [];
    _scrollPositions = {};
    _initialized = false;
    _destroyed = true;

    // 清空 DOM 缓存
    for (var key in _els) {
      if (_els.hasOwnProperty(key)) {
        if (Array.isArray(_els[key])) {
          _els[key] = [];
        } else {
          _els[key] = null;
        }
      }
    }

    console.log('[Wizard] 已销毁。');
  }

  // ──────────────────────────── 内部事件处理 ────────────────────────────

  /**
   * 键盘快捷键处理
   */
  function _onKeyDown(e) {
    // 仅在非输入框中响应
    var tag = e.target.tagName;
    if (tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT') { return; }

    switch (e.key) {
      case 'ArrowRight':
      case 'ArrowDown':
        e.preventDefault();
        nextStep();
        break;
      case 'ArrowLeft':
      case 'ArrowUp':
        e.preventDefault();
        prevStep();
        break;
      case 'Home':
        e.preventDefault();
        goToStep(0);
        break;
      case 'End':
        e.preventDefault();
        goToStep(_totalSteps - 1);
        break;
    }
  }

  /**
   * 指示器点击处理（事件委托）
   */
  function _onIndicatorClick(e) {
    var item = e.target.closest('.wizard-indicator-item');
    if (!item) { return; }

    var index = parseInt(item.getAttribute('data-indicator-index'), 10);
    if (!isNaN(index)) {
      goToStep(index);
    }
  }

  // ──────────────────────────── 自动初始化 ────────────────────────────

  document.addEventListener('DOMContentLoaded', function () {
    if (DEFAULTS.autoInit) {
      var steps = document.querySelectorAll('.wizard-step');
      if (steps.length > 0) {
        init();
      }
    }
  });

  // ──────────────────────────── 暴露全局 API ────────────────────────────

  global.Wizard = {
    init: init,
    next: nextStep,
    prev: prevStep,
    goToStep: goToStep,
    getCurrentStep: getCurrentStep,
    getTotalSteps: getTotalSteps,
    updateProgress: updateProgress,
    updateIndicator: updateIndicator,
    updateNav: updateNav,
    generate: generate,
    destroy: destroy
  };

})(window);
