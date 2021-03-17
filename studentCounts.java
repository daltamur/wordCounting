public class studentCounts {
    int wordCount;
    int contributionCount;

    //each student has a word and comment count
    public studentCounts(int wordCount, int contributionCount){
        this.wordCount=wordCount;
        this.contributionCount=contributionCount;
    }

    //increase word count by the amount given
    public void increaseWordCount(int newCount){
        wordCount=wordCount+newCount;
    }

    //increase comment count by one each time
    public void increaseContributionCount(){
        contributionCount=contributionCount+1;
    }

    public int getWordCount(){
        return wordCount;
    }

    public int getContributionCount(){
        return contributionCount;
    }
}
