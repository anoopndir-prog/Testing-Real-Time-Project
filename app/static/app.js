// Vanilla JS, no framework: styled drag-and-drop dropzones around native
// <input type=file> elements, crop-adjustable photo attach slots (seal
// identification photo + Section 5's per-sample Before/After photos), and
// add/remove rows for the Final Test Report's Distribution List.

(function () {
  "use strict";

  // Wires one dropzone's file-picked/drag-drop behavior. Used both for the
  // zones present at page load and for zones cloned in later from
  // <template> (Section 5's dynamically-added Before/After photo slots),
  // so it takes the zone element directly rather than querying the page.
  function wireDropzone(zone) {
    var input = zone.querySelector('input[type="file"]');
    if (!input) return null;

    var icon = zone.querySelector("[data-dz-icon]");
    var main = zone.querySelector("[data-dz-main]");
    var sub = zone.querySelector("[data-dz-sub]");
    var isExcel = input.accept.indexOf(".xlsx") !== -1 && input.accept.indexOf(".pdf") === -1;
    var isImage = input.accept.indexOf("image/") !== -1;

    function describe(files) {
      if (!files || files.length === 0) return;
      if (icon) icon.textContent = isImage ? "📷" : isExcel ? "📊" : "📎";
      if (files.length === 1) {
        if (main) main.textContent = files[0].name;
        if (sub) sub.textContent = "Attached successfully";
      } else {
        var names = [];
        for (var i = 0; i < Math.min(files.length, 3); i++) names.push(files[i].name);
        if (main) main.textContent = files.length + " files attached";
        if (sub) sub.textContent = names.join(", ") + (files.length > 3 ? "..." : "");
      }
    }

    input.addEventListener("change", function () {
      describe(input.files);
    });

    ["dragenter", "dragover"].forEach(function (evt) {
      zone.addEventListener(evt, function (e) {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.add("dragover");
      });
    });

    ["dragleave", "drop"].forEach(function (evt) {
      zone.addEventListener(evt, function (e) {
        e.preventDefault();
        e.stopPropagation();
        zone.classList.remove("dragover");
      });
    });

    zone.addEventListener("drop", function (e) {
      var dropped = e.dataTransfer && e.dataTransfer.files;
      if (!dropped || dropped.length === 0) return;
      if (!input.multiple && dropped.length > 1) {
        var single = new DataTransfer();
        single.items.add(dropped[0]);
        input.files = single.files;
      } else {
        input.files = dropped;
      }
      describe(input.files);
      input.dispatchEvent(new Event("change"));
    });

    return input;
  }

  function initDropzones() {
    document.querySelectorAll("[data-dropzone]").forEach(wireDropzone);
  }

  // One photo attach-and-crop slot: picking a file previews it, asks the
  // server for an auto-suggested crop box (background-subtraction
  // heuristic in app/seal_photo.py), then lets the user drag the box's
  // corners before it's baked into the report. `root` is the slot's
  // wrapping element (`[data-photo-cropper]`) - every lookup is scoped to
  // it so the same function works for the single static section 3.2 slot
  // and for however many Before/After slots Section 5 ends up with. The
  // box is always tracked in the ORIGINAL photo's pixel space
  // (`state.box`); `render()` is the only place that converts it to
  // on-screen pixels, so it stays correct across window resizes.
  function attachPhotoCropper(root, suggestUrl) {
    var zone = root.querySelector("[data-dropzone]");
    var input = zone && wireDropzone(zone);
    var container = root.querySelector("[data-seal-photo-crop]");
    var img = root.querySelector("[data-seal-photo-img]");
    var boxEl = root.querySelector("[data-seal-photo-box]");
    if (!input || !container || !img || !boxEl) return null;

    var fieldLeft = root.querySelector('[data-seal-photo-field="left"]');
    var fieldTop = root.querySelector('[data-seal-photo-field="top"]');
    var fieldRight = root.querySelector('[data-seal-photo-field="right"]');
    var fieldBottom = root.querySelector('[data-seal-photo-field="bottom"]');
    var redetectBtn = root.querySelector("[data-seal-photo-redetect]");
    var fullBtn = root.querySelector("[data-seal-photo-full]");

    var state = { naturalWidth: 0, naturalHeight: 0, box: null };

    function clamp(value, min, max) {
      return Math.min(Math.max(value, min), max);
    }

    function render() {
      if (!state.naturalWidth || !img.clientWidth) return;
      var scaleX = img.clientWidth / state.naturalWidth;
      var scaleY = img.clientHeight / state.naturalHeight;
      var left = state.box.left * scaleX;
      var top = state.box.top * scaleY;
      var right = state.box.right * scaleX;
      var bottom = state.box.bottom * scaleY;
      boxEl.style.left = left + "px";
      boxEl.style.top = top + "px";
      boxEl.style.width = Math.max(1, right - left) + "px";
      boxEl.style.height = Math.max(1, bottom - top) + "px";

      if (fieldLeft) fieldLeft.value = Math.round(state.box.left);
      if (fieldTop) fieldTop.value = Math.round(state.box.top);
      if (fieldRight) fieldRight.value = Math.round(state.box.right);
      if (fieldBottom) fieldBottom.value = Math.round(state.box.bottom);
    }

    function setFullBox() {
      state.box = { left: 0, top: 0, right: state.naturalWidth, bottom: state.naturalHeight };
      render();
    }

    function fetchSuggestion() {
      var file = input.files && input.files[0];
      if (!file || !suggestUrl) return;
      var formData = new FormData();
      formData.append("photo", file);
      fetch(suggestUrl, { method: "POST", body: formData })
        .then(function (resp) { return resp.ok ? resp.json() : null; })
        .then(function (data) {
          if (data && data.box && data.box.length === 4) {
            state.box = { left: data.box[0], top: data.box[1], right: data.box[2], bottom: data.box[3] };
            render();
          }
        })
        .catch(function () {
          // Auto-detect unavailable (offline/server error) - the full-photo
          // box set on load is still there for the user to drag by hand.
        });
    }

    input.addEventListener("change", function () {
      var file = input.files && input.files[0];
      if (!file) {
        container.hidden = true;
        return;
      }
      var reader = new FileReader();
      reader.onload = function (e) {
        img.onload = function () {
          state.naturalWidth = img.naturalWidth;
          state.naturalHeight = img.naturalHeight;
          container.hidden = false;
          setFullBox();
          fetchSuggestion();
        };
        img.src = e.target.result;
      };
      reader.readAsDataURL(file);
    });

    function startDrag(handleName, startClientX, startClientY) {
      var startBox = {
        left: state.box.left, top: state.box.top,
        right: state.box.right, bottom: state.box.bottom,
      };
      var minSize = Math.max(4, Math.min(state.naturalWidth, state.naturalHeight) * 0.01);

      function onMove(e) {
        var scaleX = state.naturalWidth / img.clientWidth;
        var scaleY = state.naturalHeight / img.clientHeight;
        var dx = (e.clientX - startClientX) * scaleX;
        var dy = (e.clientY - startClientY) * scaleY;
        var left = startBox.left, top = startBox.top, right = startBox.right, bottom = startBox.bottom;

        if (handleName === "move") {
          var w = startBox.right - startBox.left;
          var h = startBox.bottom - startBox.top;
          left = clamp(startBox.left + dx, 0, state.naturalWidth - w);
          top = clamp(startBox.top + dy, 0, state.naturalHeight - h);
          right = left + w;
          bottom = top + h;
        } else {
          if (handleName.indexOf("w") !== -1) left = clamp(startBox.left + dx, 0, right - minSize);
          if (handleName.indexOf("e") !== -1) right = clamp(startBox.right + dx, left + minSize, state.naturalWidth);
          if (handleName.indexOf("n") !== -1) top = clamp(startBox.top + dy, 0, bottom - minSize);
          if (handleName.indexOf("s") !== -1) bottom = clamp(startBox.bottom + dy, top + minSize, state.naturalHeight);
        }

        state.box = { left: left, top: top, right: right, bottom: bottom };
        render();
      }

      function onUp() {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
      }

      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
    }

    boxEl.querySelectorAll("[data-handle]").forEach(function (handle) {
      handle.addEventListener("mousedown", function (e) {
        e.preventDefault();
        e.stopPropagation();
        startDrag(handle.getAttribute("data-handle"), e.clientX, e.clientY);
      });
    });

    boxEl.addEventListener("mousedown", function (e) {
      e.preventDefault();
      startDrag("move", e.clientX, e.clientY);
    });

    if (redetectBtn) redetectBtn.addEventListener("click", fetchSuggestion);
    if (fullBtn) fullBtn.addEventListener("click", setFullBox);
    window.addEventListener("resize", render);

    return { input: input };
  }

  // Wires the photo-cropper slot(s) already present in the page at load
  // (currently just the section 3.2 seal photo). Section 5's Before/After
  // slots are wired individually as they're created - see
  // initSamplePhotos().
  function initStaticPhotoCroppers() {
    var form = document.querySelector("form[data-seal-photo-suggest-url]");
    if (!form) return;
    var suggestUrl = form.getAttribute("data-seal-photo-suggest-url");
    form.querySelectorAll(":scope > [data-photo-cropper]").forEach(function (root) {
      attachPhotoCropper(root, suggestUrl);
    });
  }

  // Section 5's per-sample Before/After Test photos. The sample count and
  // labels aren't known until the Project Spec + Test Inspection sheet
  // have been parsed server-side (same detection the report generator
  // itself uses for Sections 2/4.3/5, so the numbering always matches),
  // so this re-detects and rebuilds the photo slots whenever those inputs
  // (or the monitoring sheets) change.
  function initSamplePhotos() {
    var form = document.querySelector("form[data-detect-samples-url]");
    var statusEl = document.querySelector("[data-sample-photos-status]");
    var container = document.querySelector("[data-sample-photos-container]");
    var rowTemplate = document.querySelector("[data-sample-photo-row-template]");
    var slotTemplate = document.querySelector("[data-photo-slot-template]");
    var projectSpecInput = document.getElementById("project_spec");
    var inspectionInput = document.getElementById("test_inspection");
    var monitoringInput = document.getElementById("monitoring_sheets");
    if (!form || !statusEl || !container || !rowTemplate || !slotTemplate || !projectSpecInput || !inspectionInput) {
      return;
    }

    var detectUrl = form.getAttribute("data-detect-samples-url");
    var suggestUrl = form.getAttribute("data-seal-photo-suggest-url");
    var detectSeq = 0;
    var debounceTimer = null;

    function clearSlots() {
      container.innerHTML = "";
    }

    function buildSlot(label, phase) {
      var clone = slotTemplate.content.cloneNode(true);
      var root = clone.querySelector("[data-photo-cropper]");
      var input = clone.querySelector("[data-seal-photo-input]");
      var main = clone.querySelector("[data-slot-main]");
      var fields = clone.querySelectorAll("[data-seal-photo-field]");

      var baseName = "samplephoto_" + label.idx + "_" + phase;
      input.name = baseName;
      main.textContent = "Attach " + (phase === "before" ? "Before" : "After") + " Test photo";
      fields.forEach(function (field) {
        field.name = baseName + "_" + field.getAttribute("data-seal-photo-field");
      });

      // attachPhotoCropper needs the node already in the document (layout
      // reads like clientWidth need it) - caller appends `clone` first,
      // then wires it via attachPhotoCropper(root, ...).
      return { fragment: clone, root: root, input: input };
    }

    function buildRow(idx, label) {
      var rowClone = rowTemplate.content.cloneNode(true);
      var row = rowClone.querySelector(".sample-photo-row");
      var title = rowClone.querySelector("[data-row-title]");
      var columns = rowClone.querySelector("[data-row-columns]");
      title.textContent = label;

      var hiddenLabel = document.createElement("input");
      hiddenLabel.type = "hidden";
      hiddenLabel.name = "samplephoto_" + idx + "_label";
      hiddenLabel.value = label;
      row.appendChild(hiddenLabel);

      var before = buildSlot({ idx: idx }, "before");
      var after = buildSlot({ idx: idx }, "after");
      columns.appendChild(before.fragment);
      columns.appendChild(after.fragment);
      container.appendChild(rowClone);

      // Before-Test caption: free text describing what the photo shows
      // (e.g. "Shaft, Bore") - varies per sample/photo, so it's typed in
      // rather than auto-filled, unlike the After-Test caption which the
      // server derives from this sample's own remarks.
      var captionLabel = document.createElement("div");
      captionLabel.className = "seal-photo-hint";
      captionLabel.style.margin = "6px 0 2px";
      captionLabel.textContent = "What does this show? (optional, e.g. Shaft, Bore)";
      var captionInput = document.createElement("input");
      captionInput.type = "text";
      captionInput.name = "samplephoto_" + idx + "_before_caption";
      captionInput.className = "sample-photo-caption-input";
      before.root.appendChild(captionLabel);
      before.root.appendChild(captionInput);

      attachPhotoCropper(before.root, suggestUrl);
      attachPhotoCropper(after.root, suggestUrl);
    }

    function runDetection() {
      var seq = ++detectSeq;
      clearSlots();

      var projectSpecFile = projectSpecInput.files && projectSpecInput.files[0];
      var inspectionFile = inspectionInput.files && inspectionInput.files[0];
      if (!projectSpecFile || !inspectionFile) {
        statusEl.textContent = "Attach the Project Specification and Test Inspection sheet above to detect samples.";
        return;
      }

      statusEl.textContent = "Detecting samples…";
      var formData = new FormData();
      formData.append("project_spec", projectSpecFile);
      formData.append("test_inspection", inspectionFile);
      if (monitoringInput && monitoringInput.files) {
        for (var i = 0; i < monitoringInput.files.length; i++) {
          formData.append("monitoring_sheets", monitoringInput.files[i]);
        }
      }

      fetch(detectUrl, { method: "POST", body: formData })
        .then(function (resp) { return resp.json().then(function (data) { return { ok: resp.ok, data: data }; }); })
        .then(function (result) {
          if (seq !== detectSeq) return; // inputs changed again mid-request
          if (!result.ok || result.data.error) {
            statusEl.textContent = "Could not detect samples: " + (result.data.error || "unknown error");
            return;
          }
          var labels = result.data.labels || [];
          if (labels.length === 0) {
            statusEl.textContent = "No samples detected.";
            return;
          }
          statusEl.textContent = "Detected " + labels.length + " sample" + (labels.length !== 1 ? "s" : "") +
            " — attach Before/After photos below (optional).";
          labels.forEach(function (label, idx) { buildRow(idx, label); });
        })
        .catch(function () {
          if (seq !== detectSeq) return;
          statusEl.textContent = "Could not detect samples (network error).";
        });
    }

    function scheduleDetection() {
      if (debounceTimer) clearTimeout(debounceTimer);
      debounceTimer = setTimeout(runDetection, 400);
    }

    [projectSpecInput, inspectionInput, monitoringInput].forEach(function (input) {
      if (input) input.addEventListener("change", scheduleDetection);
    });
  }

  function initDistributionList() {
    var body = document.querySelector("[data-distribution-body]");
    var template = document.querySelector("[data-distribution-row-template]");
    var addButton = document.querySelector("[data-add-distribution-row]");
    if (!body || !template || !addButton) return;

    addButton.addEventListener("click", function () {
      var clone = template.content.cloneNode(true);
      body.appendChild(clone);
    });

    body.addEventListener("click", function (e) {
      var button = e.target.closest("[data-remove-row]");
      if (!button) return;
      var row = button.closest("tr");
      if (row) row.remove();
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    initDropzones();
    initStaticPhotoCroppers();
    initSamplePhotos();
    initDistributionList();
  });
})();
