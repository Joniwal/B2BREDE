/* Combobox de filtros com múltipla seleção, sem dependências externas. */
(() => {
  "use strict";
  const instances = new WeakMap();
  const resolve = (target) => typeof target === "string" ? document.getElementById(target) : target;
  const selectedValues = (select) => select
    ? Array.from(select.selectedOptions || []).map((option) => option.value).filter(Boolean)
    : [];

  function update(instance) {
    const selected = selectedValues(instance.select);
    let text = instance.select.dataset.allLabel || "Todos";
    if (selected.length === 1) {
      text = Array.from(instance.select.options).find((option) => option.value === selected[0])?.textContent || selected[0];
    } else if (selected.length > 1) {
      text = `${selected.length} selecionados`;
    }
    instance.button.querySelector("span").textContent = text;
    instance.panel.querySelectorAll('input[type="checkbox"]').forEach((checkbox) => {
      checkbox.checked = selected.includes(checkbox.value);
    });
  }

  function renderOptions(instance) {
    const options = Array.from(instance.select.options).filter((option) => option.value);
    instance.panel.replaceChildren();
    if (!options.length) {
      const empty = document.createElement("div");
      empty.className = "rb-multi-empty";
      empty.textContent = "Nenhuma opção disponível";
      instance.panel.appendChild(empty);
    }
    options.forEach((option) => {
      const label = document.createElement("label");
      label.className = "rb-multi-option";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = option.value;
      checkbox.checked = option.selected;
      const text = document.createElement("span");
      text.textContent = option.textContent;
      label.append(checkbox, text);
      instance.panel.appendChild(label);
      checkbox.addEventListener("change", () => {
        option.selected = checkbox.checked;
        update(instance);
        instance.select.dispatchEvent(new Event("change", { bubbles: true }));
      });
    });
    update(instance);
  }

  function enhance(select) {
    if (!select || instances.has(select)) return instances.get(select);
    select.multiple = true;
    select.classList.add("rb-multi-native");
    const root = document.createElement("div");
    root.className = "rb-multi";
    root.innerHTML = '<button class="rb-multi-button" type="button" aria-haspopup="listbox" aria-expanded="false"><span></span><i class="bi bi-chevron-down" aria-hidden="true"></i></button><div class="rb-multi-panel" role="listbox" aria-multiselectable="true"></div>';
    select.insertAdjacentElement("afterend", root);
    const instance = { select, root, button: root.querySelector("button"), panel: root.querySelector(".rb-multi-panel") };
    instances.set(select, instance);
    instance.button.addEventListener("click", (event) => {
      event.stopPropagation();
      document.querySelectorAll(".rb-multi.open").forEach((element) => { if (element !== root) element.classList.remove("open"); });
      root.classList.toggle("open");
      instance.button.setAttribute("aria-expanded", String(root.classList.contains("open")));
    });
    instance.panel.addEventListener("click", (event) => event.stopPropagation());
    new MutationObserver(() => renderOptions(instance)).observe(select, { childList: true });
    renderOptions(instance);
    return instance;
  }

  function init(root = document) {
    root.querySelectorAll?.("select[data-multi-select]").forEach(enhance);
  }

  function clear(target, notify = false) {
    const select = resolve(target);
    if (!select) return;
    Array.from(select.options).forEach((option) => { option.selected = false; });
    if (instances.has(select)) update(instances.get(select));
    if (notify) select.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function setValues(target, values, notify = false) {
    const select = resolve(target);
    if (!select) return;
    const wanted = new Set((Array.isArray(values) ? values : [values]).filter(Boolean).map(String));
    Array.from(select.options).forEach((option) => { option.selected = wanted.has(option.value); });
    if (instances.has(select)) update(instances.get(select));
    if (notify) select.dispatchEvent(new Event("change", { bubbles: true }));
  }

  document.addEventListener("click", () => document.querySelectorAll(".rb-multi.open").forEach((element) => element.classList.remove("open")));
  document.addEventListener("DOMContentLoaded", () => init());
  window.rbMultiSelect = { init, values: (target) => selectedValues(resolve(target)), clear, setValues };
})();
