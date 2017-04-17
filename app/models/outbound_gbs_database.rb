class OutboundGbsDatabase < ApplicationRecord
  def self.update_analysis
    @query = OutboundGbsDatabase.pluck('*')
    
  end
  
end
