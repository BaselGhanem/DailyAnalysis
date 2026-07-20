(() => {
    const labels = {
        latest: `طلبيات آخر يوم`,
        ytd: `من أول السنة لليوم`,
        month: `هذا الشهر`
    };

    const rangeFor = mode => {
        const anchor = meta.latestDate;
        if (!anchor) return null;

        const [year, month] = anchor.split(`-`);

        if (mode === `latest`) {
            return { from: anchor, to: anchor };
        }

        if (mode === `ytd`) {
            return { from: `${year}-01-01`, to: anchor };
        }

        return { from: `${year}-${month}-01`, to: anchor };
    };

    const updateCaptions = () => {
        if (!meta.latestDate) return;

        const latest = rangeFor(`latest`);
        const ytd = rangeFor(`ytd`);
        const month = rangeFor(`month`);

        $(`#latestDayCaption`).textContent = latest.to;
        $(`#yearToDateCaption`).textContent = `${ytd.from} ← ${ytd.to}`;
        $(`#monthCaption`).textContent = `${month.from} ← ${month.to}`;
    };

    const syncState = () => {
        if (!meta.latestDate) return;

        const from = $(`#from`).value;
        const to = $(`#to`).value;
        let active = ``;

        $$(`[data-quick-range]`).forEach(button => {
            const range = rangeFor(button.dataset.quickRange);
            const matches = from === range.from && to === range.to;

            button.classList.toggle(`active`, matches);
            if (matches) active = button.dataset.quickRange;
        });

        $(`#quickRangeStatus`).textContent = active ? labels[active] : `فترة مخصصة`;
    };

    const applyRange = (mode, silent = false) => {
        const range = rangeFor(mode);
        if (!range) return;

        $(`#from`).value = range.from;
        $(`#to`).value = range.to;
        apply();
        syncState();

        if (!silent) {
            showToast(`تم تطبيق: ${labels[mode]}`);
        }
    };

    $$(`[data-quick-range]`).forEach(button => {
        button.addEventListener(`click`, () => applyRange(button.dataset.quickRange));
    });

    [$(`#from`), $(`#to`)].forEach(input => {
        input.addEventListener(`change`, syncState);
    });

    $(`#reset`).addEventListener(`click`, () => {
        requestAnimationFrame(syncState);
    });

    const dashboard = $(`#dashboard`);
    const observer = new MutationObserver(() => {
        if (!dashboard.classList.contains(`hidden`) && meta.latestDate) {
            updateCaptions();
            applyRange(`month`, true);
            observer.disconnect();
        }
    });

    observer.observe(dashboard, { attributes: true, attributeFilter: [`class`] });
})();