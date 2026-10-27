/*
 * SPDX-FileCopyrightText: 2026 ElixirOffice contributors
 * SPDX-License-Identifier: AGPL-3.0-only
 */
define([], function () {
    'use strict';
    return Object.freeze({
        placeholder:'Ask Elixir… e.g. “Create a lesson outline”, “Shorten this slide”',
        footer:'Context sent: bounded slide text, speaker notes and drawing positions. The assistant can read the deck.',
        scope:'Selected slides, or whole deck',
        scopeDeck:'Whole deck ({count} slides)',
        scopeSlides:'Selected slides ({count})',
        read:'Reading slide text and speaker notes',
        summarize:'Summarizing the deck text',
        create:'Added {n} slides',
        rewrite:'Rewrote {n} paragraphs',
        layout:'Applied a layout to {n} slides',
        theme:'Applied an embedded theme to {n} slides',
        align:'Aligned {n} shapes',
        notes:'Updated speaker notes',
        rollbackUnverified:'The action failed. Cancellation could not be verified. Review your presentation and use Undo if needed.',
        staleDiscard:'The presentation changed after this action. Use Undo to review later changes first; no changes were discarded.'
    });
});
