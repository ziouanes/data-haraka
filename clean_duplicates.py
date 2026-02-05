import pandas as pd

input_file = r"c:\Users\HP\Desktop\data haraka\Rslt_Mvt_Enseignant_Prim2025.xlsx"
output_file = r"c:\Users\HP\Desktop\data haraka\Rslt_Mvt_Enseignant_Prim2025_cleaned.xlsx"

def fix_garbled_text(text):
	# Mapping based on your example; add more pairs as you discover them.
	mapping = {
		# --- Aleph & Lam Variants ---
		'΍': 'ا',
		'ϟ': 'ل',
		'Ϡ': 'ل',  # Lam medial
		't': 'ل',  # Lam final (visual guess based on 'Swahel')
		'Ύ': 'ا',  # Aleph medial
		'P': 'ا',  # Aleph
		'!': 'إ',  # Aleph with Hamza under
		
		# --- Basic Letters ---
		'Ο': 'ج',
		'ϭ': 'و',
		'ϋ': 'ع',
		'ό': 'ع',  # Ain medial
		'Δ': 'ة',  # Ta Marbuta
		'Γ': 'ة',  # Ta Marbuta 2
		'~': 'س',
		'Α': 'ب',
		'Ώ': 'ب',  # Ba final
		'λ': 'ص',
		'ι': 'ص',  # Sad final
		'΢': 'ح',
		'Η': 'ت',
		'Ε': 'ت',  # Ta final (open)
		'η': 'ش',
		'ε': 'ش',  # Sheen final
		'ρ': 'ط',
		'ί': 'ز',  # Zay
		'ϧ': 'ن',
		'ϥ': 'ن',  # Noon final
		'ϲ': 'ي',  # Ya final
		'΋': 'ئ',  # Ya with Hamza
		'ο': 'ض',
		'ę': 'م',  # Meem final
		'ϫ': 'ه',
		'Λ': 'ث',
		'κ': 'ك',
		'Χ': 'خ',
		'ϔ': 'ف',  # Fa medial
		'ϑ': 'ف',  # Fa final
		'θ': 'ق',  # Qaf (often theta)
		'Ω': 'د',
		'ϣ': 'م',
		'Σ': 'ح',
		' ': ' ',
		'έ': 'ر',
		'ϳ': 'ي',
		'Ϙ': 'ق',
		'ϓ': 'ف',
		'ϕ': 'ق',
		'Ν': 'ج',
		'ϱ': 'ي',
		'ϗ': 'ق',
		'΃': 'أ',
		'اϷ': 'الأ',
		'ϐ': 'غ',
		'ϼ': 'لا',
		'α': 'س',
		'ϛ': 'ك',
		'ϻ': 'لا',
        ':': 'غ',
		'Ϋ': 'ذ',
		'ϖ': 'ق',
		' ': ' ',
		' ': ' ',
		
    
	
	}

	decoded_chars = []
	for char in text:
		decoded_chars.append(mapping.get(char, char))

	decoded_text = "".join(decoded_chars)
	return decoded_text[::-1]

print("Loading file...")
df = pd.read_excel(input_file, dtype=str)

print("Fixing garbled text...")
for column in df.columns:
	df[column] = df[column].apply(
		lambda value: fix_garbled_text(value) if isinstance(value, str) else value
	)

# Remove full-row duplicates (exact match across all columns)
print("Removing duplicates...")
original_count = len(df)
df = df.drop_duplicates(keep="first")

# Save the result
print("Saving cleaned file...")
df.to_excel(output_file, index=False)

print(f"Done. Output saved to: {output_file}")
print(f"Total rows after cleaning: {len(df)}")
print(f"Duplicates removed: {original_count - len(df)}")
