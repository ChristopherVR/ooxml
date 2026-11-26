/**
 * The Symbols gallery of the Equation tab, grouped into Word's categories. Each category is the
 * characters in gallery order; a symbol is inserted as itself.
 */

export type EquationSymbolCategoryId =
	| 'basicMath'
	| 'greekLetters'
	| 'letterLike'
	| 'operators'
	| 'arrows'
	| 'negatedRelations'
	| 'scripts'
	| 'geometry';

export interface EquationSymbolCategory {
	id: EquationSymbolCategoryId;
	/** English name; the editors translate it. */
	name: string;
	symbols: readonly string[];
}

const chars = (text: string) => Array.from(text.replace(/\s/g, ''));

export const EQUATION_SYMBOLS: readonly EquationSymbolCategory[] = [
	{
		id: 'basicMath',
		name: 'Basic Math',
		symbols: chars(
			'± ∞ = ≠ ~ × ÷ ! ∝ < ≪ > ≫ ≤ ≥ ∓ ≅ ≈ ≡ ∀ ∁ ∂ √ ∛ ∜ ∪ ∩ ∅ % ° ℉ ℃ ∆ ∇ ∃ ∄ ∈ ∋ ← ↑ → ↓ ↔ ∴ + − ¬ α β γ δ ε ϵ θ ϑ μ π ρ σ τ φ ω ∗ ∙ ⋮ ⋯ ⋰ ⋱ ℵ ℶ ∎',
		),
	},
	{
		id: 'greekLetters',
		name: 'Greek Letters',
		symbols: chars(
			'α β γ δ ε ϵ ζ η θ ϑ ι κ λ μ ν ξ ο π ϖ ρ ϱ σ ς τ υ φ ϕ χ ψ ω Α Β Γ Δ Ε Ζ Η Θ Ι Κ Λ Μ Ν Ξ Ο Π Ρ Σ Τ Υ Φ Χ Ψ Ω',
		),
	},
	{
		id: 'letterLike',
		name: 'Letter-Like Symbols',
		symbols: chars(
			'∀ ∁ ℂ ∂ ð ℇ Ϝ Ⅎ ℊ ℋ ℌ ℍ ℎ ℏ ı ℑ ȷ ℒ ℓ ℕ ℘ ℙ ℚ ℛ ℜ ℝ ℤ K Å ℬ ℭ ℯ ℰ ℱ ℳ ℴ ℵ ℶ ℷ ℸ',
		),
	},
	{
		id: 'operators',
		name: 'Operators',
		symbols: chars(
			'+ − ± ∓ ∗ ∘ ∙ × ÷ ⊕ ⊖ ⊗ ⊘ ⊙ ∧ ∨ ∩ ∪ ⊎ ⊓ ⊔ ∖ ≀ ≔ ≕ ≝ ≜ ≞ ≟ ≡ ∼ ≃ ≅ ≈ ≍ ≐ ∝ ≺ ≻ ≼ ≽ ⊂ ⊃ ⊆ ⊇ ⊏ ⊐ ⊑ ⊒ ∈ ∋ ⊢ ⊣ ⊨ ∣ ∥ ⊥ ⋈ ⋉ ⋊',
		),
	},
	{
		id: 'arrows',
		name: 'Arrows',
		symbols: chars(
			'← → ↑ ↓ ↔ ↕ ⇐ ⇒ ⇑ ⇓ ⇔ ⇕ ⟵ ⟶ ⟷ ⟸ ⟹ ⟺ ↗ ↖ ↘ ↙ ↚ ↛ ↮ ⇍ ⇏ ⇎ ↤ ↦ ↩ ↪ ↼ ↽ ⇀ ⇁ ↿ ↾ ⇃ ⇂ ⇋ ⇌ ⇄ ⇆ ⇈ ⇊ ⇉ ⇇ ↞ ↠ ↶ ↷ ↺ ↻ ⊸',
		),
	},
	{
		id: 'negatedRelations',
		name: 'Negated Relations',
		symbols: chars('≠ ≮ ≯ ≰ ≱ ≢ ≁ ≄ ≉ ≇ ≭ ⊀ ⊁ ⋠ ⋡ ∉ ∌ ⊄ ⊅ ⊈ ⊉ ⋢ ⋣ ⊊ ⊋ ∤ ∦ ⊬ ⊭ ⊮ ⊯ ⋪ ⋫ ⋬ ⋭'),
	},
	{
		id: 'scripts',
		name: 'Scripts',
		symbols: chars(
			'𝒜 ℬ 𝒞 𝒟 ℰ ℱ 𝒢 ℋ ℐ 𝒥 𝒦 ℒ ℳ 𝒩 𝒪 𝒫 𝒬 ℛ 𝒮 𝒯 𝒰 𝒱 𝒲 𝒳 𝒴 𝒵 𝔄 𝔅 ℭ 𝔇 𝔈 𝔉 𝔊 ℌ ℑ 𝔍 𝔎 𝔏 𝔐 𝔑 𝔒 𝔓 𝔔 ℜ 𝔖 𝔗 𝔘 𝔙 𝔚 𝔛 𝔜 ℨ 𝔸 𝔹 ℂ 𝔻 𝔼 𝔽 𝔾 ℍ 𝕀 𝕁 𝕂 𝕃 𝕄 ℕ 𝕆 ℙ ℚ ℝ 𝕊 𝕋 𝕌 𝕍 𝕎 𝕏 𝕐 ℤ',
		),
	},
	{
		id: 'geometry',
		name: 'Geometry',
		symbols: chars('∟ ∠ ∡ ∢ ⊾ ⊿ ⋕ ⊥ ∤ ∣ ∥ ∦ ∶ ∷ ∴ ∵ ∎ □ ▭ ▱ △ ▷ ◁ ◇ ○ ◊ ⟂ ≅ ∼ ≃'),
	},
];
