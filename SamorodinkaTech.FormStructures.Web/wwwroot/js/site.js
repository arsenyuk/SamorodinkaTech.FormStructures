// Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.

document.addEventListener('click', (e) => {
	const target = e.target;
	if (!(target instanceof HTMLElement)) {
		return;
	}

	// If the user clicked on an interactive element inside the row, let it work normally.
	if (target.closest('a, button, input, label, select, textarea')) {
		return;
	}

	const row = target.closest('.js-row-link');
	if (!(row instanceof HTMLElement)) {
		return;
	}

	const href = row.getAttribute('data-href');
	if (!href) {
		return;
	}

	window.location.href = href;
});
