export type ValidateDataDecisionKey = "fixMissing" | "fixDuplicates" | "fixBadType";
export type ValidateDataDetectKey = "missing" | "duplicates" | "badType";
export interface ValidateDataQuestionDef {
  id: "validate_data:fix_missing" | "validate_data:fix_duplicates" | "validate_data:fix_bad_type";
  decisionKey: ValidateDataDecisionKey;
  question: string;
  detectKey: ValidateDataDetectKey;
}

export const VALIDATE_DATA_QUESTIONS: ValidateDataQuestionDef[] = [
  {
    id: "validate_data:fix_missing",
    decisionKey: "fixMissing",
    question: "Supprimer les lignes avec données manquantes ?",
    detectKey: "missing",
  },
  {
    id: "validate_data:fix_duplicates",
    decisionKey: "fixDuplicates",
    question: "Supprimer les lignes en doublon ?",
    detectKey: "duplicates",
  },
  {
    id: "validate_data:fix_bad_type",
    decisionKey: "fixBadType",
    question: "Typecaster les valeurs (si conversion non ambiguë) ?",
    detectKey: "badType",
  },
];

export const VALIDATE_DATA_CONFIRM_CHOICES = [
  { id: "yes", label: "Oui" },
  { id: "no", label: "Non" },
];

export const VALIDATE_DATA_CONFIRM_IDS: Set<string> = new Set(
  VALIDATE_DATA_QUESTIONS.map((question) => question.id)
);
