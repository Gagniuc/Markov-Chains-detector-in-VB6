# :fleur_de_lis: Markov Chains detector in VB6

The current VB6 application is a detector that uses two models, a model "+" that is associated with what we are looking for, and a model "-" that is associated with the background. Both models are represented by a transition matrix that is calculated or trained by using two sequences of observations. Namely, a sequence of observations that is known to belong to a region of interest (model "+") and a sequence of observations that may represent either a random sequence or a sequence other than the sequence "+". Once the model sequences have been used to construct the transition matrices for the two models, they are merged into a single matrix, namely into a log-likelihood matrix (LLM). The log-likelihood matrix represents "the memory", a kind of signature that can be used in some detections. But how? A scanner can use this LLM to search for model-like "+" regions inside a longer sequence called <i>z</i> (the target). To search for such regions of interest, sliding windows are used. The content of a sliding window is examined by [verifying each transition with the values from the LLM](https://figshare.com/articles/figure/Local_score_computation_by_using_the_LLM_pdf/19205124). Once a transition is associated with a value, it is added to the previous result until all transitions in the content of the sliding window are verified. This [results in a main score for each sliding window over <i>z</i>](https://figshare.com/articles/figure/An_experiment_for_understanding_scores_pdf/19205067). The positive scores (red) indicate the regions that resemble the "+" model, and the negative scores indicate that the content of the sliding window is different from the "+" model.

[This version in JS](https://gagniuc.github.io/Markov-Chains-scanner/) can also be of use: [Markov Chains detector in Javascript](https://github.com/Gagniuc/Markov-Chains-scanner)

<kbd><img src="https://github.com/Gagniuc/Markov-Chains-detector-in-VB6/blob/main/screenshot/Markov%20Chains%20detector%20in%20VB6%20(2).gif" /></kbd>

<kbd><img src="https://github.com/Gagniuc/Markov-Chains-detector-in-VB6/blob/main/screenshot/Markov%20Chains%20detector%20in%20VB6%20(3).gif" /></kbd>

# References

- <i>Paul A. Gagniuc. Algorithms in Bioinformatics: Theory and Implementation. John Wiley & Sons, Hoboken, NJ, USA, 2021, ISBN: 9781119697961.</i>

