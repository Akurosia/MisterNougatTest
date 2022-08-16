module Jekyll

  module Humanize
    ##
    # This is a port of the Django app `humanize` which adds a "human touch"
    # to data. Given that Jekyll produces static sites, some of the original
    # methods do not make logical sense (e.g. naturaltime).
    #
    # Source code can be viewed here:
    # https://github.com/django/django
    #
    # Copyright (c) Django Software Foundation and individual contributors.
    # All rights reserved.

    ####################
    #  PUBLIC METHODS  #
    ####################

    def ordinal(value, flag=nil)
      ##
      # Converts an integer to its ordinal as a string. 1 is '1st', 2 is '2nd',
      # 3 is '3rd', etc. Works for any integer.
      #
      # Usage:
      # {{ somenum }} >>> 3
      # {{ somenum | ordinal }} >>> '3rd'
      # {{ somenum | ordinal: "super" }} >>> '3<sup>rd</sup>'

      begin
        value = value.to_i
        flag.to_s.downcase!
      rescue Exception => e
        puts "#{e.class} #{e}"
        return value
      end

      suffix = ""
      suffixes = ["th", "st", "nd", "rd", "th", "th", "th", "th", "th", "th"]
      unless [11, 12, 13].include? value % 100 then
        suffix = suffixes[value % 10]
      else
        suffix = suffixes[0]
      end

      unless flag and flag == "super"
        return "#{value}%s" % suffix
      else
        return "#{value}<sup>%s</sup>" % suffix
      end

    end

    def intcomma(value, delimiter=",")
      ##
      # Converts an integer to a string containing commas every three digits.
      # For example, 3000 becomes '3,000' and 45000 becomes '45,000'.
      # Optionally supports a delimiter override for commas.
      #
      # Usage:
      # {{ post.content | number_of_words }} >>> 12345
      # {{ post.content | number_of_words | intcomma }} >>> '12,345'
      # {{ post.content | number_of_words | intcomma: '.' }} >>> '12.345'

      begin
        orig = value.to_s
        delimiter = delimiter.to_s
      rescue Exception => e
        puts "#{e.class} #{e}"
        return value
      end

      copy = orig.strip
      copy = orig.gsub(/^(-?\d+)(\d{3})/, "\\1#{delimiter}\\2")
      orig == copy ? copy : intcomma(copy, delimiter)
    end

    INTWORD_HELPERS = [
      [6, "million"],
      [9, "billion"],
      [12, "trillion"],
      [15, "quadrillion"],
      [18, "quintillion"],
      [21, "sextillion"],
      [24, "septillion"],
      [27, "octillion"],
      [30, "nonillion"],
      [33, "decillion"],
      [100, "googol"],
    ]

    def intword(value)
      ##
      # Converts a large integer to a friendly text representation. Works best
      # for numbers over 1 million. For example, 1000000 becomes '1.0 million',
      # 1200000 becomes '1.2 million' and 1200000000 becomes '1.2 billion'.
      #
      # Usage:
      # {{ largenum }} >>> 1200000
      # {{ largenum | intword }} >>> '1.2 million'

      begin
        value = value.to_i
      rescue Exception => e
        puts "#{e.class} #{e}"
        return value
      end

      if value < 1000000
        return value
      end

      for exponent, text in INTWORD_HELPERS
        large_number = 10 ** exponent

        if value < large_number * 1000
          return "%#{value}.1f #{text}" % (value / large_number.to_f)
        end

      end

      return value
    end

    def filesize(value)
      ##
      # For filesize values in bytes, returns the number rounded to 3
      # decimal places with the correct suffix.
      #
      # Usage:
      # {{ bytes }} >>> 123456789
      # {{ bytes | filesize }} >>> 117.738 MB
      filesize_tb = 1099511627776.0
      filesize_gb = 1073741824.0
      filesize_mb = 1048576.0
      filesize_kb = 1024.0

      begin
        value = value.to_f
      rescue Exception => e
        puts "#{e.class} #{e}"
        return value
      end

      if value >= filesize_tb
        return "%s TB" % (value / filesize_tb).to_f.round(3)
      elsif value >= filesize_gb
        return "%s GB" % (value / filesize_gb).to_f.round(3)
      elsif value >= filesize_mb
        return "%s MB" % (value / filesize_mb).to_f.round(3)
      elsif value >= filesize_kb
        return "%s KB" % (value / filesize_kb).to_f.round(0)
      elsif value == 1
        return "1 byte"
      else
        return "%s bytes" % value.to_f.round(0)
      end

    end

  end

end

Liquid::Template.register_filter(Jekyll::Humanize)
