import type { BlobSample } from "./types";

const csvContent = `question,selected_answer,correct_answer,time,result,explanation
"<p>Testv 2</p>",B,D,0s,Wrong,ddddd
"True or False: The sum of probabilities of all possible outcomes of a random experiment is always less than 1.",False,False,0s,Correct,"The sum of probabilities of all possible outcomes of a random experiment is always equal to 1."
"A box contains 10 defective and 90 non-defective items. If one item is selected at random, what is the probability that it is non-defective?",1/9,9/10,0s,Wrong,"The probability of selecting a non-defective item is 90/100 = 9/10."
`;

export const csvBlobSample: BlobSample = {
  id: "csv",
  label: "CSV Blob",
  description: "CSV content passed as a Blob plus filename metadata.",
  createSource: () => ({
    blob: new Blob([csvContent], { type: "text/csv" }),
    fileName: "mid-term-report.csv",
  }),
};
