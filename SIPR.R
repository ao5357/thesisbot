
SIPR <- function(metadata,batch_directory){
#Load XML library
library(XML)

#Some variables to use later
SIP.directory <- batch_directory

#Change data from factors to strings
if(class(metadata$path)=="factor") metadata <- data.frame(lapply(metadata, as.character),stringsAsFactors=FALSE)


#Check if files exist
for(i in 1:length(metadata$path)){
	FileNotFound <- "Cannot locate file"
	FileFound <- "Found"
	exists <- file.exists(metadata$path[i])	
	if(exists == FALSE) print(paste(FileNotFound,metadata$path[i])) else print(paste(FileFound,metadata$path[i]))	
}

#Create SIP folder structure and contents files and copy PDFs/files
dir.create(SIP.directory)

for(i in 1:length(metadata$path)){
#Create folders		
	slash ="\\"
	item <- unlist(strsplit(metadata$path[i],"\\",fixed=T))
	item.name <- unlist(strsplit(item[length(item)],".",fixed=T))
	name <- item.name[1]
	dir.create(paste(SIP.directory,slash,name,sep=""))

#Create contents files
	contents.name <- "contents"
	contents <- file(paste(SIP.directory,slash,name,slash,contents.name,sep=""))
	writeLines(item[length(item)],contents)
	close(contents)

#Copy PDFs/files into folders
	file.copy(metadata$path[i],paste(SIP.directory,slash,name,sep=""))
}

#Create XML files for metadata
	metadata.names <- strsplit(names(metadata),".",fixed=T)
	metadata.names[[length(metadata.names)]] <- NULL
	
	#Get namespaces from metadata element prefixes -- "dc" is considered default
	namespaces <- NULL
	for(i in 1:length(metadata.names)){
		namespaces[1] <- "dc"
		element <- unlist(metadata.names[[i]])
		if(element[1] == namespaces[1]) NULL else namespaces[length(namespaces)+1] <- element[1]
		}
		
	#Put values into XML
	metadata.nopath <- metadata[,-(ncol(metadata))]
	for(i in 1:nrow(metadata.nopath)){
		
		#Create XML nodes
		dc <- newXMLNode("dublin_core",attrs=c(schema="dc"))
		cat(saveXML(dc))	
		if(length(namespaces) > 1){
		alt <- newXMLNode(namespaces[2],attrs=c(schema=namespaces[2])) 
		cat(saveXML(alt)) }
		
		for(j in 1:ncol(metadata.nopath)){
		if(metadata.nopath[i,j] != ""){	
			prefix.tag <- unlist(strsplit(unlist(names(metadata.nopath[j])),".",fixed=T))
			prefix <- prefix.tag[1]
			tag <- prefix.tag[2]
			subtag <- prefix.tag[3]
		
			repeated <- grep("||", metadata.nopath[i,j])
			
				if(length(repeated) == 1){
			
					delimited <- unlist(strsplit(metadata.nopath[i,j],split="||",fixed=T))
			
					for(r in 1:length(delimited)){
						
					if(as.name(prefix)=="dc") addChildren(dc,newXMLNode("dcvalue",delimited[r],attrs=c(element=tag,qualifier=if(is.na(subtag)==FALSE) subtag else "none"))) else addChildren(alt,newXMLNode("dcvalue",delimited[r],attrs=c(element=tag,qualifier=if(is.na(subtag)==FALSE) subtag else "none")))
			
					}
				}
				
				else
		
		if(as.name(prefix) == "dc") addChildren(dc,newXMLNode("dcvalue",metadata.nopath[i,j],attrs=c(element=tag, qualifier=if(is.na(subtag)==FALSE) subtag else "none"))) else addChildren(alt,newXMLNode("dcvalue",metadata.nopath[i,j], attrs=c(element=tag, qualifier=if(is.na(subtag)==FALSE) subtag else "none")))
		}
		}
		
		#get item names for saving
		item <- unlist(strsplit(metadata$path[i],"\\",fixed=T))
		item.name <- unlist(strsplit(item[length(item)],".",fixed=T))
		name <- item.name[1]
		
		saveXML(dc,paste(SIP.directory,slash,name,slash,"dublin_core.xml",sep=""))	
		if(exists(as.character(substitute(alt))) == TRUE){
		saveXML(alt,paste(SIP.directory,slash,name,slash,namespaces[2],".xml",sep=""	))}
	}	
print("Process completed.")	
}				