FROM r-base
COPY . /home/docker
WORKDIR /home/docker
RUN apt-get update
RUN apt-get install curl -qq
RUN apt-get install libcurl4-openssl-dev -qq
RUN apt-get install libssl-dev -qq
RUN apt-get install libxml2-dev -qq
RUN Rscript -e 'install.packages("caTools", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("bizdays", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("functional", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("collections", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("gtools", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("stringi", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("AzureAuth", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("AzureGraph", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("httr", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("httpuv", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("jsonlite", verbose=FALSE, quiet=TRUE)'
RUN Rscript -e 'install.packages("tidyverse", verbose=FALSE, quiet=TRUE)'
CMD ["ifconfig -a"]
CMD ["Rscript", "MailReaderBatch.R"]