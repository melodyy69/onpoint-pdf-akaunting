(() => {
  const dropZone = document.querySelector("#drop-zone");
  const pdfInput = document.querySelector("#pdf-input");
  const fileLabel = document.querySelector("#file-label");

  if (dropZone && pdfInput && fileLabel) {
    const showFile = () => {
      const file = pdfInput.files && pdfInput.files[0];
      fileLabel.textContent = file ? file.name : "Перетащите PDF сюда";
    };

    ["dragenter", "dragover"].forEach((eventName) => {
      dropZone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropZone.classList.add("is-dragging");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      dropZone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropZone.classList.remove("is-dragging");
      });
    });

    dropZone.addEventListener("drop", (event) => {
      const files = event.dataTransfer && event.dataTransfer.files;
      if (files && files.length) {
        pdfInput.files = files;
        showFile();
      }
    });
    pdfInput.addEventListener("change", showFile);
  }

  const reviewApp = document.querySelector("#review-app");
  if (!reviewApp) return;

  const dataElement = document.querySelector("#unknown-transactions");
  const groups = dataElement ? JSON.parse(dataElement.textContent || "[]") : [];
  const categoryOptionsElement = document.querySelector("#category-options");
  const savedCategories = categoryOptionsElement
    ? JSON.parse(categoryOptionsElement.textContent || "[]")
    : [];
  const categories = new Set(savedCategories);
  const initialCount = groups.length;
  let currentIndex = 0;

  const panels = document.querySelector("#review-panels");
  const complete = document.querySelector("#review-complete");
  const form = document.querySelector("#category-form");
  const input = document.querySelector("#category-input");
  const error = document.querySelector("#category-error");
  const saveButton = document.querySelector("#save-category");
  const label = document.querySelector("#progress-label");
  const bar = document.querySelector("#progress-bar");
  const categoryPicker = document.querySelector("#category-picker");
  const categorySlider = document.querySelector("#category-slider");
  const categoryPrev = document.querySelector("#category-prev");
  const categoryNext = document.querySelector("#category-next");

  const normalizedCategory = (value) => value.trim().toLocaleLowerCase();

  const updateCategorySelection = () => {
    if (!categorySlider) return;
    const currentValue = normalizedCategory(input ? input.value : "");
    categorySlider.querySelectorAll(".category-chip").forEach((button) => {
      const isSelected = normalizedCategory(button.dataset.category || "") === currentValue;
      button.classList.toggle("is-selected", isSelected);
      button.setAttribute("aria-pressed", String(isSelected));
    });
  };

  const renderCategorySlider = () => {
    if (!categorySlider || !categoryPicker) return;
    const categoryNames = [...categories]
      .filter((category) => typeof category === "string" && category.trim())
      .sort((left, right) => left.localeCompare(right, ["ru", "en"], { sensitivity: "base" }));

    categoryPicker.hidden = categoryNames.length === 0;
    categorySlider.replaceChildren();
    categoryNames.forEach((categoryName) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "category-chip";
      button.dataset.category = categoryName;
      button.textContent = categoryName;
      button.setAttribute("aria-pressed", "false");
      button.addEventListener("click", () => {
        if (input) {
          input.value = categoryName;
          input.focus();
        }
        updateCategorySelection();
      });
      categorySlider.append(button);
    });
    updateCategorySelection();
  };

  const scrollCategories = (direction) => {
    if (!categorySlider) return;
    const distance = Math.max(220, categorySlider.clientWidth * 0.75);
    categorySlider.scrollBy({ left: direction * distance, behavior: "smooth" });
  };

  if (categoryPrev) categoryPrev.addEventListener("click", () => scrollCategories(-1));
  if (categoryNext) categoryNext.addEventListener("click", () => scrollCategories(1));
  if (input) input.addEventListener("input", updateCategorySelection);

  const setText = (selector, value) => {
    const element = document.querySelector(selector);
    if (element) element.textContent = value;
  };

  const renderCurrent = () => {
    const group = groups[currentIndex];
    const completed = currentIndex >= groups.length;
    if (panels) panels.hidden = completed;
    if (complete) complete.hidden = !completed;

    const doneCount = Math.min(currentIndex, initialCount);
    const percent = initialCount === 0 ? 100 : Math.round((doneCount / initialCount) * 100);
    if (bar) bar.style.width = `${percent}%`;
    if (label) label.textContent = completed ? "Все правила сохранены" : `Осталось правил: ${groups.length - currentIndex}`;
    if (completed || !group) return;

    setText("#transaction-type", group.transaction_type === "income" ? "Income" : "Expense");
    setText("#transaction-description", group.merchant_text || group.description);
    setText("#transaction-full-description", group.description);
    setText("#transaction-dates", group.dates.join(", "));
    setText("#transaction-amounts", group.amounts.map((value) => `$${value}`).join(", "));
    setText("#transaction-count", String(group.count));
    if (input) {
      input.value = "";
      input.focus();
    }
    updateCategorySelection();
    if (error) error.textContent = "";
  };

  if (form) {
    form.addEventListener("submit", async (event) => {
      event.preventDefault();
      const group = groups[currentIndex];
      const categoryName = input ? input.value.trim() : "";
      if (!group || !categoryName) {
        if (error) error.textContent = "Введите название категории.";
        return;
      }

      if (saveButton) {
        saveButton.disabled = true;
        saveButton.textContent = "Сохраняю…";
      }
      if (error) error.textContent = "";

      try {
        const csrf = document.querySelector('meta[name="csrf-token"]')?.content || "";
        const response = await fetch(reviewApp.dataset.saveUrl, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "X-CSRF-Token": csrf,
          },
          body: JSON.stringify({
            rule_key: group.rule_key,
            category_name: categoryName,
          }),
        });
        const result = await response.json();
        if (!response.ok) throw new Error(result.error || "Не удалось сохранить категорию.");
        categories.add(result.category_name || categoryName);
        renderCategorySlider();
        currentIndex += 1;
        renderCurrent();
      } catch (requestError) {
        if (error) error.textContent = requestError.message;
      } finally {
        if (saveButton) {
          saveButton.disabled = false;
          saveButton.textContent = "Сохранить и продолжить";
        }
      }
    });
  }

  renderCategorySlider();
  renderCurrent();
})();
